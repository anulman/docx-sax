const requiredEventTypes = ['package:start', 'part:start', 'relationship', 'element', 'text', 'end', 'part:end', 'package:end'];

function fail(message) {
  throw new Error(message);
}

function assertObject(value, label) {
  if (value === null || typeof value !== 'object' || Array.isArray(value)) {
    fail(`${label} must be an object`);
  }
}

function assertString(event, property) {
  if (typeof event[property] !== 'string') {
    fail(`${event.type} event ordinal ${event.ordinal} must include string ${property}`);
  }
}

function assertNumber(event, property) {
  if (typeof event[property] !== 'number') {
    fail(`${event.type} event ordinal ${event.ordinal} must include number ${property}`);
  }
}

function assertBoolean(event, property) {
  if (typeof event[property] !== 'boolean') {
    fail(`${event.type} event ordinal ${event.ordinal} must include boolean ${property}`);
  }
}

function assertXmlShape(event) {
  assertString(event, 'partUri');
  assertString(event, 'name');
  assertString(event, 'localName');
  assertString(event, 'prefix');
  assertString(event, 'namespaceUri');
  assertNumber(event, 'depth');
  assertString(event, 'path');
}

export function eventKind(event) {
  if (event.type === 'package' || event.type === 'part') {
    return `${event.type}:${event.phase}`;
  }

  return event.type;
}

export function collectEventShapeSummary(events) {
  const seen = new Set();
  const keysByKind = new Map();

  for (const event of events) {
    seen.add(eventKind(event));
    const key = eventKind(event);
    if (!keysByKind.has(key)) {
      keysByKind.set(key, Object.keys(event).sort());
    }
  }

  return {
    count: events.length,
    seen: [...seen].sort(),
    keysByKind: Object.fromEntries([...keysByKind.entries()].sort(([a], [b]) => a.localeCompare(b))),
  };
}

export function assertCommonDocxSaxEventModel(events) {
  if (!Array.isArray(events)) {
    fail('events must be an array for test validation');
  }

  if (events.length === 0) {
    fail('expected at least one DocxSax event');
  }

  events.forEach((event, index) => {
    assertObject(event, `event[${index}]`);
    assertString(event, 'type');
    assertNumber(event, 'ordinal');

    if (event.ordinal !== index) {
      fail(`event ordinal ${event.ordinal} should equal stream index ${index}`);
    }

    switch (event.type) {
      case 'package':
        if (event.phase !== 'start' && event.phase !== 'end') {
          fail(`package event ordinal ${event.ordinal} must have start/end phase`);
        }
        break;
      case 'part':
        if (event.phase !== 'start' && event.phase !== 'end') {
          fail(`part event ordinal ${event.ordinal} must have start/end phase`);
        }
        assertString(event, 'uri');
        assertString(event, 'contentType');
        assertString(event, 'relationshipType');
        break;
      case 'relationship':
        assertString(event, 'sourceUri');
        assertString(event, 'id');
        assertString(event, 'relationshipType');
        assertString(event, 'targetUri');
        assertBoolean(event, 'isExternal');
        break;
      case 'element':
        assertXmlShape(event);
        assertBoolean(event, 'isEmptyElement');
        if (!Array.isArray(event.attributes)) {
          fail(`element event ordinal ${event.ordinal} must include attributes array`);
        }
        for (const attribute of event.attributes) {
          assertObject(attribute, `attribute on event ordinal ${event.ordinal}`);
          assertString(attribute, 'name');
          assertString(attribute, 'localName');
          assertString(attribute, 'prefix');
          assertString(attribute, 'namespaceUri');
          assertString(attribute, 'value');
        }
        break;
      case 'text':
        assertString(event, 'partUri');
        assertString(event, 'text');
        assertNumber(event, 'depth');
        assertString(event, 'path');
        assertBoolean(event, 'isWhitespace');
        break;
      case 'end':
        assertXmlShape(event);
        break;
      case 'diagnostic':
        assertString(event, 'message');
        if (event.partUri !== undefined && typeof event.partUri !== 'string' && event.partUri !== null) {
          fail(`diagnostic event ordinal ${event.ordinal} partUri must be a string or null`);
        }
        break;
      default:
        fail(`unexpected event type ${event.type} at ordinal ${event.ordinal}`);
    }
  });

  const seen = new Set(events.map(eventKind));
  const missing = requiredEventTypes.filter((type) => !seen.has(type));
  if (missing.length > 0) {
    fail(`missing expected common event kinds: ${missing.join(', ')}`);
  }

  if (!events.some((event) => event.type === 'part' && event.uri === '/word/document.xml')) {
    fail('expected public fixture to include /word/document.xml part');
  }

  if (!events.some((event) => event.type === 'text' && event.text === 'Hello DOCX SAX')) {
    fail('expected public fixture to include Hello DOCX SAX text event');
  }

  return collectEventShapeSummary(events);
}
