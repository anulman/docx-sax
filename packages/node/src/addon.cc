#include <node_api.h>

#include <dlfcn.h>

#include <cassert>
#include <cstring>
#include <memory>
#include <mutex>
#include <string>
#include <unordered_map>
#include <utility>
#include <vector>

namespace {

using BatchCallback = int (*)(const unsigned char*, int, void*);
using ParseFileJsonBatches = int (*)(const char*, int, BatchCallback, void*);

constexpr int kStatusSuccess = 0;
constexpr int kStatusInvalidArgument = 2;
constexpr int kStatusParseFailure = 3;
constexpr int kStatusCallbackFailure = 4;

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
    auto loaded_symbol = reinterpret_cast<ParseFileJsonBatches>(dlsym(handle, "docx_sax_parse_file_json_batches"));
    const char* message = dlerror();
    if (message != nullptr || loaded_symbol == nullptr) {
      error = message == nullptr ? "missing docx_sax_parse_file_json_batches export" : message;
      return false;
    }

    // .NET Native AOT libraries are process-lifetime components in this bridge;
    // dlclose after managed runtime initialization can crash on process teardown.
    Cache().emplace(path, CachedLibrary{handle, loaded_symbol});
    symbol = loaded_symbol;
    return true;
  }

  std::string path;
  ParseFileJsonBatches symbol = nullptr;
  std::string error;

 private:
  struct CachedLibrary {
    void* handle;
    ParseFileJsonBatches symbol;
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

struct ParseWork {
  napi_env env = nullptr;
  napi_async_work work = nullptr;
  napi_deferred deferred = nullptr;
  std::string input_path;
  std::string library_path;
  int batch_size = 128;
  int native_status = kStatusSuccess;
  std::string error;
  std::vector<std::string> batches;
};

int OnBatch(const unsigned char* data, int length, void* user_data) {
  if (data == nullptr || length < 0 || user_data == nullptr) {
    return 1;
  }

  auto* parse_work = static_cast<ParseWork*>(user_data);
  try {
    parse_work->batches.emplace_back(reinterpret_cast<const char*>(data), static_cast<size_t>(length));
    return 0;
  } catch (...) {
    return 1;
  }
}

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

void ExecuteParse(napi_env /*env*/, void* data) {
  auto* parse_work = static_cast<ParseWork*>(data);

  NativeLibrary library(parse_work->library_path);
  if (!library.Open()) {
    parse_work->native_status = -1;
    parse_work->error = "Unable to load DocxSax native library at '" + parse_work->library_path + "': " + library.error;
    return;
  }

  parse_work->native_status = library.symbol(
      parse_work->input_path.c_str(),
      parse_work->batch_size,
      OnBatch,
      parse_work);

  if (parse_work->native_status != kStatusSuccess) {
    parse_work->error = StatusMessage(parse_work->native_status);
  }
}

void CompleteParse(napi_env env, napi_status status, void* data) {
  std::unique_ptr<ParseWork> parse_work(static_cast<ParseWork*>(data));

  if (status != napi_ok || parse_work->native_status != kStatusSuccess) {
    napi_value error;
    const std::string message = status == napi_ok
        ? parse_work->error
        : "DocxSax native parser async work was cancelled or failed";
    Check(env, napi_create_string_utf8(env, message.c_str(), message.size(), &error));
    Check(env, napi_reject_deferred(env, parse_work->deferred, error));
  } else {
    napi_value array;
    Check(env, napi_create_array_with_length(env, parse_work->batches.size(), &array));

    for (size_t i = 0; i < parse_work->batches.size(); ++i) {
      const std::string& batch = parse_work->batches[i];
      napi_value value;
      Check(env, napi_create_string_utf8(env, batch.c_str(), batch.size(), &value));
      Check(env, napi_set_element(env, array, static_cast<uint32_t>(i), value));
    }

    Check(env, napi_resolve_deferred(env, parse_work->deferred, array));
  }

  Check(env, napi_delete_async_work(env, parse_work->work));
}

napi_value ParseFileBatchesJson(napi_env env, napi_callback_info info) {
  size_t argc = 3;
  napi_value args[3];
  Check(env, napi_get_cb_info(env, info, &argc, args, nullptr, nullptr));

  if (argc < 3) {
    napi_throw_type_error(env, nullptr, "parseFileBatchesJson(path, batchSize, nativeLibraryPath) requires 3 arguments");
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
  if (argc >= 2) {
    napi_valuetype batch_type;
    Check(env, napi_typeof(env, args[1], &batch_type));
    if (batch_type == napi_number) {
      Check(env, napi_get_value_int32(env, args[1], &batch_size));
    }
  }

  auto parse_work = std::make_unique<ParseWork>();
  parse_work->env = env;
  parse_work->input_path = ReadString(env, args[0]);
  parse_work->library_path = ReadString(env, args[2]);
  parse_work->batch_size = batch_size <= 0 ? 128 : batch_size;

  napi_value promise;
  Check(env, napi_create_promise(env, &parse_work->deferred, &promise));

  napi_value resource_name;
  Check(env, napi_create_string_utf8(env, "DocxSaxParse", NAPI_AUTO_LENGTH, &resource_name));
  Check(env, napi_create_async_work(
      env,
      nullptr,
      resource_name,
      ExecuteParse,
      CompleteParse,
      parse_work.get(),
      &parse_work->work));
  Check(env, napi_queue_async_work(env, parse_work->work));

  parse_work.release();
  return promise;
}

napi_value Init(napi_env env, napi_value exports) {
  napi_value function;
  Check(env, napi_create_function(env, "parseFileBatchesJson", NAPI_AUTO_LENGTH, ParseFileBatchesJson, nullptr, &function));
  Check(env, napi_set_named_property(env, exports, "parseFileBatchesJson", function));
  return exports;
}

}  // namespace

NAPI_MODULE(NODE_GYP_MODULE_NAME, Init)
