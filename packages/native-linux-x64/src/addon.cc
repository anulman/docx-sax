#include <node_api.h>

#include <dlfcn.h>

#include <cassert>
#include <condition_variable>
#include <cstdint>
#include <cstring>
#include <deque>
#include <memory>
#include <mutex>
#include <string>
#include <unordered_map>
#include <utility>
#include <vector>

namespace {

struct NativeStringView {
  const unsigned char* data;
  int32_t length;
};

struct NativeAttributeView {
  NativeStringView name;
  NativeStringView local_name;
  NativeStringView prefix;
  NativeStringView namespace_uri;
  NativeStringView value;
};

struct NativeDocxEventView {
  int32_t kind;
  int64_t ordinal;
  NativeStringView phase;
  NativeStringView uri;
  NativeStringView content_type;
  NativeStringView relationship_type;
  NativeStringView source_uri;
  NativeStringView id;
  NativeStringView target_uri;
  int32_t is_external;
  NativeStringView part_uri;
  NativeStringView name;
  NativeStringView local_name;
  NativeStringView prefix;
  NativeStringView namespace_uri;
  int32_t depth;
  NativeStringView path;
  int32_t is_empty_element;
  const NativeAttributeView* attributes;
  int32_t attribute_count;
  NativeStringView text;
  int32_t is_whitespace;
  NativeStringView message;
};

struct NativeAttribute {
  std::string name;
  std::string local_name;
  std::string prefix;
  std::string namespace_uri;
  std::string value;
};

struct NativeEvent {
  int32_t kind = 0;
  int64_t ordinal = 0;
  std::string phase;
  std::string uri;
  std::string content_type;
  std::string relationship_type;
  std::string source_uri;
  std::string id;
  std::string target_uri;
  bool is_external = false;
  std::string part_uri;
  std::string name;
  std::string local_name;
  std::string prefix;
  std::string namespace_uri;
  int32_t depth = 0;
  std::string path;
  bool is_empty_element = false;
  std::vector<NativeAttribute> attributes;
  std::string text;
  bool is_whitespace = false;
  std::string message;
};

using EventCallback = int (*)(const NativeDocxEventView*, void*);
using ParseFileEvents = int (*)(const char*, int, EventCallback, void*);

constexpr int kStatusSuccess = 0;
constexpr int kStatusInvalidArgument = 2;
constexpr int kStatusParseFailure = 3;
constexpr int kStatusCallbackFailure = 4;
constexpr size_t kMaxQueuedBatches = 4;

void Check(napi_env env, napi_status status) {
  assert(status == napi_ok);
  (void)env;
  (void)status;
}

std::string ReadString(napi_env env, napi_value value) {
  size_t length = 0;
  Check(env, napi_get_value_string_utf8(env, value, nullptr, 0, &length));

  std::vector<char> buffer(length + 1);
  size_t copied = 0;
  Check(env, napi_get_value_string_utf8(env, value, buffer.data(), buffer.size(), &copied));
  return std::string(buffer.data(), copied);
}

std::string CopyString(NativeStringView value) {
  if (value.data == nullptr || value.length < 0) {
    return {};
  }

  return std::string(reinterpret_cast<const char*>(value.data), static_cast<size_t>(value.length));
}

struct NativeLibrary {
  explicit NativeLibrary(std::string path) : path(std::move(path)) {}

  bool Open() {
    std::lock_guard<std::mutex> guard(CacheMutex());
    auto cached = Cache().find(path);
    if (cached != Cache().end()) {
      symbol = cached->second.symbol;
      return true;
    }

    dlerror();
    void* handle = dlopen(path.c_str(), RTLD_NOW | RTLD_LOCAL);
    if (handle == nullptr) {
      const char* message = dlerror();
      error = message == nullptr ? "failed to load native library" : message;
      return false;
    }

    dlerror();
    auto loaded_symbol = reinterpret_cast<ParseFileEvents>(dlsym(handle, "docx_sax_parse_file_events"));
    const char* message = dlerror();
    if (message != nullptr || loaded_symbol == nullptr) {
      error = message == nullptr ? "missing docx_sax_parse_file_events export" : message;
      // Symbol validation failed before the .NET Native AOT entrypoint could run,
      // so this handle is not a process-lifetime managed runtime yet.
      dlclose(handle);
      return false;
    }

    // Once the required export is validated, treat .NET Native AOT libraries as
    // process-lifetime components; dlclose after managed runtime initialization
    // can crash on process teardown.
    Cache().emplace(path, CachedLibrary{handle, loaded_symbol});
    symbol = loaded_symbol;
    return true;
  }

  std::string path;
  ParseFileEvents symbol = nullptr;
  std::string error;

 private:
  struct CachedLibrary {
    void* handle;
    ParseFileEvents symbol;
  };

  static std::mutex& CacheMutex() {
    static std::mutex mutex;
    return mutex;
  }

  static std::unordered_map<std::string, CachedLibrary>& Cache() {
    static std::unordered_map<std::string, CachedLibrary> cache;
    return cache;
  }
};

std::string StatusMessage(int status) {
  switch (status) {
    case kStatusInvalidArgument:
      return "invalid argument passed to DocxSax native parser";
    case kStatusParseFailure:
      return "DocxSax native parser failed to parse input DOCX";
    case kStatusCallbackFailure:
      return "DocxSax native parser callback failed while collecting a batch";
    default:
      return "DocxSax native parser returned status " + std::to_string(status);
  }
}

struct PendingNext {
  napi_deferred deferred = nullptr;
};

struct ParseStream {
  napi_env env = nullptr;
  napi_async_work work = nullptr;
  napi_threadsafe_function notifier = nullptr;
  int64_t id = 0;
  std::string input_path;
  std::string library_path;
  int batch_size = 128;
  int native_status = kStatusSuccess;
  std::string error;

  std::mutex mutex;
  std::condition_variable queue_changed;
  std::deque<std::vector<NativeEvent>> queued_batches;
  std::vector<NativeEvent> current_batch;
  std::deque<PendingNext> pending_nexts;
  bool done = false;
  bool disposed = false;
};

std::mutex& StreamsMutex() {
  static std::mutex mutex;
  return mutex;
}

std::unordered_map<int64_t, std::shared_ptr<ParseStream>>& Streams() {
  static std::unordered_map<int64_t, std::shared_ptr<ParseStream>> streams;
  return streams;
}

int64_t NextStreamId() {
  static int64_t next = 1;
  std::lock_guard<std::mutex> guard(StreamsMutex());
  return next++;
}

std::shared_ptr<ParseStream> FindStream(int64_t id) {
  std::lock_guard<std::mutex> guard(StreamsMutex());
  auto stream = Streams().find(id);
  return stream == Streams().end() ? nullptr : stream->second;
}

void RemoveStream(int64_t id) {
  std::lock_guard<std::mutex> guard(StreamsMutex());
  Streams().erase(id);
}

void SetStringProperty(napi_env env, napi_value object, const char* name, const std::string& value) {
  napi_value string_value;
  Check(env, napi_create_string_utf8(env, value.c_str(), value.size(), &string_value));
  Check(env, napi_set_named_property(env, object, name, string_value));
}

void SetInt64Property(napi_env env, napi_value object, const char* name, int64_t value) {
  napi_value number_value;
  Check(env, napi_create_int64(env, value, &number_value));
  Check(env, napi_set_named_property(env, object, name, number_value));
}

void SetInt32Property(napi_env env, napi_value object, const char* name, int32_t value) {
  napi_value number_value;
  Check(env, napi_create_int32(env, value, &number_value));
  Check(env, napi_set_named_property(env, object, name, number_value));
}

void SetBooleanProperty(napi_env env, napi_value object, const char* name, bool value) {
  napi_value bool_value;
  Check(env, napi_get_boolean(env, value, &bool_value));
  Check(env, napi_set_named_property(env, object, name, bool_value));
}

napi_value CreateAttribute(napi_env env, const NativeAttribute& attribute) {
  napi_value object;
  Check(env, napi_create_object(env, &object));
  SetStringProperty(env, object, "name", attribute.name);
  SetStringProperty(env, object, "localName", attribute.local_name);
  SetStringProperty(env, object, "prefix", attribute.prefix);
  SetStringProperty(env, object, "namespaceUri", attribute.namespace_uri);
  SetStringProperty(env, object, "value", attribute.value);
  return object;
}

napi_value CreateAttributes(napi_env env, const std::vector<NativeAttribute>& attributes) {
  napi_value array;
  Check(env, napi_create_array_with_length(env, attributes.size(), &array));
  for (size_t index = 0; index < attributes.size(); index++) {
    Check(env, napi_set_element(env, array, static_cast<uint32_t>(index), CreateAttribute(env, attributes[index])));
  }
  return array;
}

napi_value CreateEvent(napi_env env, const NativeEvent& event) {
  napi_value object;
  Check(env, napi_create_object(env, &object));

  switch (event.kind) {
    case 0: // PackageStart
    case 1: // PackageEnd
      SetStringProperty(env, object, "type", "package");
      SetStringProperty(env, object, "phase", event.phase);
      SetInt64Property(env, object, "ordinal", event.ordinal);
      break;
    case 2: // PartStart
    case 3: // PartEnd
      SetStringProperty(env, object, "type", "part");
      SetStringProperty(env, object, "phase", event.phase);
      SetInt64Property(env, object, "ordinal", event.ordinal);
      SetStringProperty(env, object, "uri", event.uri);
      SetStringProperty(env, object, "contentType", event.content_type);
      SetStringProperty(env, object, "relationshipType", event.relationship_type);
      break;
    case 4: // Relationship
      SetStringProperty(env, object, "type", "relationship");
      SetInt64Property(env, object, "ordinal", event.ordinal);
      SetStringProperty(env, object, "sourceUri", event.source_uri);
      SetStringProperty(env, object, "id", event.id);
      SetStringProperty(env, object, "relationshipType", event.relationship_type);
      SetStringProperty(env, object, "targetUri", event.target_uri);
      SetBooleanProperty(env, object, "isExternal", event.is_external);
      break;
    case 5: // ElementStart
      SetStringProperty(env, object, "type", "element");
      SetInt64Property(env, object, "ordinal", event.ordinal);
      SetStringProperty(env, object, "partUri", event.part_uri);
      SetStringProperty(env, object, "name", event.name);
      SetStringProperty(env, object, "localName", event.local_name);
      SetStringProperty(env, object, "prefix", event.prefix);
      SetStringProperty(env, object, "namespaceUri", event.namespace_uri);
      SetInt32Property(env, object, "depth", event.depth);
      SetStringProperty(env, object, "path", event.path);
      SetBooleanProperty(env, object, "isEmptyElement", event.is_empty_element);
      Check(env, napi_set_named_property(env, object, "attributes", CreateAttributes(env, event.attributes)));
      break;
    case 6: // ElementEnd
      SetStringProperty(env, object, "type", "end");
      SetInt64Property(env, object, "ordinal", event.ordinal);
      SetStringProperty(env, object, "partUri", event.part_uri);
      SetStringProperty(env, object, "name", event.name);
      SetStringProperty(env, object, "localName", event.local_name);
      SetStringProperty(env, object, "prefix", event.prefix);
      SetStringProperty(env, object, "namespaceUri", event.namespace_uri);
      SetInt32Property(env, object, "depth", event.depth);
      SetStringProperty(env, object, "path", event.path);
      break;
    case 7: // Text
      SetStringProperty(env, object, "type", "text");
      SetInt64Property(env, object, "ordinal", event.ordinal);
      SetStringProperty(env, object, "partUri", event.part_uri);
      SetStringProperty(env, object, "text", event.text);
      SetInt32Property(env, object, "depth", event.depth);
      SetStringProperty(env, object, "path", event.path);
      SetBooleanProperty(env, object, "isWhitespace", event.is_whitespace);
      break;
    case 8: // Diagnostic
      SetStringProperty(env, object, "type", "diagnostic");
      SetInt64Property(env, object, "ordinal", event.ordinal);
      SetStringProperty(env, object, "message", event.message);
      if (!event.part_uri.empty()) {
        SetStringProperty(env, object, "partUri", event.part_uri);
      }
      break;
    default:
      SetStringProperty(env, object, "type", "diagnostic");
      SetInt64Property(env, object, "ordinal", event.ordinal);
      SetStringProperty(env, object, "message", "Unknown native DocxSax event kind");
      break;
  }

  return object;
}

napi_value CreateBatch(napi_env env, const std::vector<NativeEvent>& events) {
  napi_value array;
  Check(env, napi_create_array_with_length(env, events.size(), &array));
  for (size_t index = 0; index < events.size(); index++) {
    Check(env, napi_set_element(env, array, static_cast<uint32_t>(index), CreateEvent(env, events[index])));
  }
  return array;
}

napi_value CreateNextResult(napi_env env, bool done, const std::vector<NativeEvent>* value) {
  napi_value result;
  Check(env, napi_create_object(env, &result));

  napi_value done_value;
  Check(env, napi_get_boolean(env, done, &done_value));
  Check(env, napi_set_named_property(env, result, "done", done_value));

  if (value != nullptr) {
    Check(env, napi_set_named_property(env, result, "value", CreateBatch(env, *value)));
  }

  return result;
}

void ResolveDeferred(napi_env env, napi_deferred deferred, bool done, const std::vector<NativeEvent>* value) {
  napi_value result = CreateNextResult(env, done, value);
  Check(env, napi_resolve_deferred(env, deferred, result));
}

void RejectDeferred(napi_env env, napi_deferred deferred, const std::string& message) {
  napi_value error;
  Check(env, napi_create_string_utf8(env, message.c_str(), message.size(), &error));
  Check(env, napi_reject_deferred(env, deferred, error));
}

void DrainPending(napi_env env, ParseStream* stream) {
  struct Resolution {
    napi_deferred deferred;
    bool done;
    std::vector<NativeEvent> value;
    bool has_value;
    bool reject;
    std::string error;
  };

  std::vector<Resolution> resolutions;

  {
    std::lock_guard<std::mutex> guard(stream->mutex);
    while (!stream->pending_nexts.empty()) {
      auto deferred = stream->pending_nexts.front().deferred;

      if (!stream->queued_batches.empty()) {
        std::vector<NativeEvent> value = std::move(stream->queued_batches.front());
        stream->queued_batches.pop_front();
        stream->pending_nexts.pop_front();
        resolutions.push_back(Resolution{deferred, false, std::move(value), true, false, {}});
        stream->queue_changed.notify_all();
        continue;
      }

      if (stream->done || stream->disposed) {
        stream->pending_nexts.pop_front();
        if (stream->native_status != kStatusSuccess && !stream->disposed) {
          resolutions.push_back(Resolution{deferred, true, {}, false, true, stream->error});
        } else {
          resolutions.push_back(Resolution{deferred, true, {}, false, false, {}});
        }
        continue;
      }

      break;
    }
  }

  for (const auto& resolution : resolutions) {
    if (resolution.reject) {
      RejectDeferred(env, resolution.deferred, resolution.error);
    } else if (resolution.has_value) {
      ResolveDeferred(env, resolution.deferred, resolution.done, &resolution.value);
    } else {
      ResolveDeferred(env, resolution.deferred, resolution.done, nullptr);
    }
  }
}

void NotifyMainThread(const std::shared_ptr<ParseStream>& stream) {
  if (stream->notifier != nullptr) {
    (void)napi_call_threadsafe_function(stream->notifier, nullptr, napi_tsfn_nonblocking);
  }
}

void NotifierCallback(napi_env env, napi_value /*js_callback*/, void* context, void* /*data*/) {
  if (env == nullptr) {
    return;
  }

  auto* stream = static_cast<ParseStream*>(context);
  DrainPending(env, stream);
}

NativeEvent CopyEvent(const NativeDocxEventView& native_event) {
  NativeEvent event;
  event.kind = native_event.kind;
  event.ordinal = native_event.ordinal;
  event.phase = CopyString(native_event.phase);
  event.uri = CopyString(native_event.uri);
  event.content_type = CopyString(native_event.content_type);
  event.relationship_type = CopyString(native_event.relationship_type);
  event.source_uri = CopyString(native_event.source_uri);
  event.id = CopyString(native_event.id);
  event.target_uri = CopyString(native_event.target_uri);
  event.is_external = native_event.is_external != 0;
  event.part_uri = CopyString(native_event.part_uri);
  event.name = CopyString(native_event.name);
  event.local_name = CopyString(native_event.local_name);
  event.prefix = CopyString(native_event.prefix);
  event.namespace_uri = CopyString(native_event.namespace_uri);
  event.depth = native_event.depth;
  event.path = CopyString(native_event.path);
  event.is_empty_element = native_event.is_empty_element != 0;
  event.text = CopyString(native_event.text);
  event.is_whitespace = native_event.is_whitespace != 0;
  event.message = CopyString(native_event.message);

  if (native_event.attributes != nullptr && native_event.attribute_count > 0) {
    event.attributes.reserve(static_cast<size_t>(native_event.attribute_count));
    for (int32_t index = 0; index < native_event.attribute_count; index++) {
      const auto& attribute = native_event.attributes[index];
      event.attributes.push_back(NativeAttribute{
          CopyString(attribute.name),
          CopyString(attribute.local_name),
          CopyString(attribute.prefix),
          CopyString(attribute.namespace_uri),
          CopyString(attribute.value),
      });
    }
  }

  return event;
}

bool FlushCurrentBatch(const std::shared_ptr<ParseStream>& stream, std::unique_lock<std::mutex>& lock) {
  if (stream->current_batch.empty()) {
    return true;
  }

  stream->queue_changed.wait(lock, [&stream] {
    return stream->disposed || stream->queued_batches.size() < kMaxQueuedBatches;
  });

  if (stream->disposed) {
    return false;
  }

  stream->queued_batches.emplace_back(std::move(stream->current_batch));
  stream->current_batch.clear();
  return true;
}

int OnEvent(const NativeDocxEventView* event, void* user_data) {
  if (event == nullptr || user_data == nullptr) {
    return 1;
  }

  auto* holder = static_cast<std::shared_ptr<ParseStream>*>(user_data);
  const auto& stream = *holder;

  std::unique_lock<std::mutex> lock(stream->mutex);
  if (stream->disposed) {
    return 1;
  }

  try {
    stream->current_batch.push_back(CopyEvent(*event));
    if (stream->current_batch.size() >= static_cast<size_t>(stream->batch_size)) {
      if (!FlushCurrentBatch(stream, lock)) {
        return 1;
      }
      lock.unlock();
      NotifyMainThread(stream);
      return 0;
    }
  } catch (...) {
    return 1;
  }

  return 0;
}

void ExecuteParse(napi_env /*env*/, void* data) {
  auto* holder = static_cast<std::shared_ptr<ParseStream>*>(data);
  const auto& stream = *holder;

  NativeLibrary library(stream->library_path);
  if (!library.Open()) {
    std::lock_guard<std::mutex> guard(stream->mutex);
    stream->native_status = -1;
    stream->error = "Unable to load DocxSax native library at '" + stream->library_path + "': " + library.error;
    stream->done = true;
    return;
  }

  int status = library.symbol(
      stream->input_path.c_str(),
      stream->batch_size,
      OnEvent,
      holder);

  {
    std::unique_lock<std::mutex> lock(stream->mutex);
    if (status == kStatusSuccess && !stream->disposed && !FlushCurrentBatch(stream, lock)) {
      status = kStatusCallbackFailure;
    }
    stream->native_status = status;
    if (status != kStatusSuccess && !stream->disposed) {
      stream->error = StatusMessage(status);
    }
    stream->done = true;
  }
  stream->queue_changed.notify_all();
  NotifyMainThread(stream);
}

void CompleteParse(napi_env env, napi_status status, void* data) {
  std::unique_ptr<std::shared_ptr<ParseStream>> holder(static_cast<std::shared_ptr<ParseStream>*>(data));
  const auto& stream = *holder;

  if (status != napi_ok) {
    std::lock_guard<std::mutex> guard(stream->mutex);
    stream->native_status = -1;
    stream->error = "DocxSax native parser async work was cancelled or failed";
    stream->done = true;
  }

  stream->queue_changed.notify_all();
  NotifyMainThread(stream);
  DrainPending(env, stream.get());

  if (stream->notifier != nullptr) {
    Check(env, napi_release_threadsafe_function(stream->notifier, napi_tsfn_release));
  }
  Check(env, napi_delete_async_work(env, stream->work));
}

int32_t ReadStreamId(napi_env env, napi_value value) {
  int32_t id = 0;
  Check(env, napi_get_value_int32(env, value, &id));
  return id;
}

napi_value StartParseFileBatches(napi_env env, napi_callback_info info) {
  size_t argc = 3;
  napi_value args[3];
  Check(env, napi_get_cb_info(env, info, &argc, args, nullptr, nullptr));

  if (argc < 3) {
    napi_throw_type_error(env, nullptr, "startParseFileBatches(path, batchSize, nativeLibraryPath) requires 3 arguments");
    return nullptr;
  }

  napi_valuetype path_type;
  napi_valuetype library_type;
  Check(env, napi_typeof(env, args[0], &path_type));
  Check(env, napi_typeof(env, args[2], &library_type));
  if (path_type != napi_string || library_type != napi_string) {
    napi_throw_type_error(env, nullptr, "path and nativeLibraryPath must be strings");
    return nullptr;
  }

  int32_t batch_size = 128;
  napi_valuetype batch_type;
  Check(env, napi_typeof(env, args[1], &batch_type));
  if (batch_type == napi_number) {
    Check(env, napi_get_value_int32(env, args[1], &batch_size));
  }

  auto stream = std::make_shared<ParseStream>();
  stream->env = env;
  stream->id = NextStreamId();
  stream->input_path = ReadString(env, args[0]);
  stream->library_path = ReadString(env, args[2]);
  stream->batch_size = batch_size <= 0 ? 128 : batch_size;

  napi_value resource_name;
  Check(env, napi_create_string_utf8(env, "DocxSaxParseStream", NAPI_AUTO_LENGTH, &resource_name));
  Check(env, napi_create_threadsafe_function(
      env,
      nullptr,
      nullptr,
      resource_name,
      0,
      1,
      nullptr,
      nullptr,
      stream.get(),
      NotifierCallback,
      &stream->notifier));

  auto holder = std::make_unique<std::shared_ptr<ParseStream>>(stream);
  Check(env, napi_create_async_work(
      env,
      nullptr,
      resource_name,
      ExecuteParse,
      CompleteParse,
      holder.get(),
      &stream->work));

  {
    std::lock_guard<std::mutex> guard(StreamsMutex());
    Streams().emplace(stream->id, stream);
  }

  Check(env, napi_queue_async_work(env, stream->work));
  holder.release();

  napi_value id_value;
  Check(env, napi_create_int32(env, static_cast<int32_t>(stream->id), &id_value));
  return id_value;
}

napi_value NextBatch(napi_env env, napi_callback_info info) {
  size_t argc = 1;
  napi_value args[1];
  Check(env, napi_get_cb_info(env, info, &argc, args, nullptr, nullptr));

  if (argc < 1) {
    napi_throw_type_error(env, nullptr, "nextBatch(streamId) requires 1 argument");
    return nullptr;
  }

  auto stream = FindStream(ReadStreamId(env, args[0]));
  if (stream == nullptr) {
    napi_value promise;
    napi_deferred deferred;
    Check(env, napi_create_promise(env, &deferred, &promise));
    ResolveDeferred(env, deferred, true, nullptr);
    return promise;
  }

  napi_value promise;
  napi_deferred deferred;
  Check(env, napi_create_promise(env, &deferred, &promise));

  {
    std::lock_guard<std::mutex> guard(stream->mutex);
    stream->pending_nexts.push_back(PendingNext{deferred});
  }

  DrainPending(env, stream.get());
  return promise;
}

napi_value DisposeParse(napi_env env, napi_callback_info info) {
  size_t argc = 1;
  napi_value args[1];
  Check(env, napi_get_cb_info(env, info, &argc, args, nullptr, nullptr));

  if (argc >= 1) {
    auto id = ReadStreamId(env, args[0]);
    auto stream = FindStream(id);
    if (stream != nullptr) {
      {
        std::unique_lock<std::mutex> lock(stream->mutex);
        stream->disposed = true;
        stream->queue_changed.notify_all();
        stream->queue_changed.wait(lock, [&stream] {
          return stream->done;
        });
      }
      NotifyMainThread(stream);
      RemoveStream(id);
    }
  }

  napi_value undefined;
  Check(env, napi_get_undefined(env, &undefined));
  return undefined;
}

napi_value Init(napi_env env, napi_value exports) {
  napi_value start_function;
  Check(env, napi_create_function(env, "startParseFileBatches", NAPI_AUTO_LENGTH, StartParseFileBatches, nullptr, &start_function));
  Check(env, napi_set_named_property(env, exports, "startParseFileBatches", start_function));

  napi_value next_function;
  Check(env, napi_create_function(env, "nextBatch", NAPI_AUTO_LENGTH, NextBatch, nullptr, &next_function));
  Check(env, napi_set_named_property(env, exports, "nextBatch", next_function));

  napi_value dispose_function;
  Check(env, napi_create_function(env, "disposeParse", NAPI_AUTO_LENGTH, DisposeParse, nullptr, &dispose_function));
  Check(env, napi_set_named_property(env, exports, "disposeParse", dispose_function));

  return exports;
}

}  // namespace

NAPI_MODULE(NODE_GYP_MODULE_NAME, Init)
