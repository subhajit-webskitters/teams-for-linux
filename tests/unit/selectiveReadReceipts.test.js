"use strict";

const test = require("node:test");
const assert = require("node:assert/strict");

const {
  isReadReceiptUrl,
  extractConversationId,
  getMessageId,
  SelectiveReadReceiptsController,
} = require("../../app/browser/tools/selectiveReadReceipts");

test("isReadReceiptUrl detects consumption horizon and read receipt endpoints", () => {
  const receiptUrls = [
    "https://emea.ng.msg.teams.microsoft.com/v1/users/ME/conversations/19:chat1/consumptionhorizon",
    "https://teams.microsoft.com/api/chatsvc/emea/v1/users/ME/conversations/19:chat2/consumptionhorizons",
    "https://teams.cloud.microsoft/api/readreceipt",
    "https://teams.microsoft.com/v1/users/ME/conversations/xyz/readreceipts",
    "https://example.com/api/ConsumptionHorizon?param=1",
  ];

  for (const url of receiptUrls) {
    assert.equal(isReadReceiptUrl(url), true, `Failed for: ${url}`);
  }

  const nonReceiptUrls = [
    "https://emea.ng.msg.teams.microsoft.com/v1/users/ME/conversations/19:chat1/messages",
    "https://teams.microsoft.com/api/auth",
    "https://teams.cloud.microsoft/api/profile",
    "",
    null,
    undefined,
    12345,
  ];

  for (const url of nonReceiptUrls) {
    assert.equal(isReadReceiptUrl(url), false, `Should be false for: ${url}`);
  }
});

test("extractConversationId extracts thread IDs and conversation names", () => {
  assert.equal(
    extractConversationId(
      "https://emea.ng.msg.teams.microsoft.com/v1/users/ME/conversations/19:abc123_def456@thread.v2/consumptionhorizon"
    ),
    "19:abc123_def456@thread.v2"
  );

  assert.equal(
    extractConversationId(
      "https://teams.microsoft.com/_#/conversations/19:meeting_xyz@thread.tacv2?ctx=chat"
    ),
    "19:meeting_xyz@thread.tacv2"
  );

  assert.equal(
    extractConversationId(
      "https://teams.cloud.microsoft/v2/chat/19:channel-123@thread.skype"
    ),
    "19:channel-123@thread.skype"
  );

  assert.equal(
    extractConversationId(
      "https://teams.microsoft.com/conversations/team_general_chat/consumptionhorizon"
    ),
    "team_general_chat"
  );

  assert.equal(
    extractConversationId(
      "https://teams.microsoft.com/api/chatsvc/conversations/19%3Asomething%40thread.v2"
    ),
    "19:something@thread.v2"
  );

  assert.equal(extractConversationId("https://teams.microsoft.com/"), "default");
  assert.equal(extractConversationId(""), "default");
  assert.equal(extractConversationId(null), "default");
});

test("getMessageId extracts data-mid or id attributes", () => {
  const elWithMid = {
    nodeType: 1,
    getAttribute: (attr) => (attr === "data-mid" ? "mid-12345" : null),
  };
  assert.equal(getMessageId(elWithMid), "mid-12345");

  const elWithId = {
    nodeType: 1,
    id: "message-789",
    getAttribute: () => null,
  };
  assert.equal(getMessageId(elWithId), "message-789");

  const elEmpty = {
    nodeType: 1,
    getAttribute: () => null,
  };
  assert.equal(getMessageId(elEmpty), null);
  assert.equal(getMessageId(null), null);
});

test("isMyMessage correctly identifies own messages", () => {
  const ctrl = new SelectiveReadReceiptsController({});

  const myMsg1 = {
    nodeType: 1,
    matches: (sel) => sel.includes("fui-ChatMyMessage"),
    closest: () => null,
  };
  assert.equal(ctrl.isMyMessage(myMsg1), true);

  const myMsg2 = {
    nodeType: 1,
    matches: () => false,
    closest: (sel) => sel.includes("--mine"),
  };
  assert.equal(ctrl.isMyMessage(myMsg2), true);

  const incomingMsg = {
    nodeType: 1,
    matches: () => false,
    closest: () => null,
  };
  assert.equal(ctrl.isMyMessage(incomingMsg), false);
  assert.equal(ctrl.isMyMessage(null), false);
});

test("network interceptor buffers read receipts and releases on manual trigger", async () => {
  let originalFetchCalled = false;
  let originalFetchUrl = null;

  const mockOriginalFetch = async (input, init) => {
    originalFetchCalled = true;
    originalFetchUrl = typeof input === "string" ? input : input?.url;
    return { ok: true, status: 200 };
  };

  const mockWindow = {
    location: { hostname: "teams.microsoft.com", href: "https://teams.microsoft.com" },
    fetch: mockOriginalFetch,
  };

  const ctrl = new SelectiveReadReceiptsController({
    window: mockWindow,
    document: { getElementById: () => null, createElement: () => ({ setAttribute: () => {} }) },
  });

  ctrl.installNetworkInterceptor();

  // Test intercepting a normal fetch
  const normalRes = await mockWindow.fetch("https://teams.microsoft.com/api/profile", {
    method: "GET",
  });
  assert.equal(normalRes.ok, true);
  assert.equal(originalFetchCalled, true);
  assert.equal(ctrl.pendingCount, 0);

  // Reset flag
  originalFetchCalled = false;

  // Test intercepting a consumption horizon fetch
  const horizonUrl =
    "https://emea.ng.msg.teams.microsoft.com/v1/users/ME/conversations/19:chat-test/consumptionhorizon";
  const horizonRes = await mockWindow.fetch(horizonUrl, {
    method: "PUT",
    body: JSON.stringify({ consumptionhorizon: "12345" }),
  });

  // Interceptor should buffer it and return 200 OK without calling originalFetch
  assert.equal(horizonRes.status, 200);
  assert.equal(originalFetchCalled, false);
  assert.equal(ctrl.pendingCount, 1);

  // Releasing the read receipt manually
  const released = ctrl.sendReadReceipt("19:chat-test");
  assert.equal(released, true);
  assert.equal(originalFetchCalled, true);
  assert.equal(originalFetchUrl, horizonUrl);
  assert.equal(ctrl.pendingCount, 0);
});

test("DOM discovery attaches eye button and clicking marks seen", () => {
  let buttonCreated = null;
  const elements = [];

  class MockElement {
    constructor(tagName) {
      this.tagName = tagName;
      this.nodeType = 1;
      this.attributes = new Map();
      this.children = [];
      this.parentElement = null;
      this.previousElementSibling = null;
      this.nextSibling = null;
      this.listeners = new Map();
      this.innerHTML = "";
    }

    setAttribute(name, val) {
      this.attributes.set(name, String(val));
    }

    getAttribute(name) {
      return this.attributes.get(name) || null;
    }

    hasAttribute(name) {
      return this.attributes.has(name);
    }

    addEventListener(event, fn) {
      this.listeners.set(event, fn);
    }

    dispatchEvent(event) {
      const fn = this.listeners.get(event.type);
      if (fn) fn(event);
    }

    querySelector(selector) {
      if (selector.includes("data-tfl-receipt-ui")) {
        for (const child of this.children) {
          if (child.hasAttribute("data-tfl-receipt-ui")) return child;
        }
      }
      return null;
    }

    matches(selector) {
      if (selector.includes("chat-pane-message")) {
        return this.getAttribute("data-tid") === "chat-pane-message";
      }
      return false;
    }

    closest() {
      return null;
    }

    insertBefore(child, reference) {
      this.children.push(child);
      child.parentElement = this;
    }

    prepend(child) {
      this.children.unshift(child);
      child.parentElement = this;
    }
  }

  const mockDocument = {
    getElementById: () => null,
    createElement: (tag) => {
      const el = new MockElement(tag);
      if (tag === "button") buttonCreated = el;
      return el;
    },
    head: new MockElement("head"),
    body: new MockElement("body"),
  };

  const mockWindow = {
    location: { hostname: "teams.microsoft.com", href: "https://teams.microsoft.com" },
  };

  const ctrl = new SelectiveReadReceiptsController({
    window: mockWindow,
    document: mockDocument,
  });

  const incomingMessage = new MockElement("div");
  incomingMessage.setAttribute("data-tid", "chat-pane-message");
  incomingMessage.setAttribute("data-mid", "mid-999");

  // Run ensureEyeControl
  ctrl.ensureEyeControl(incomingMessage);

  assert.ok(buttonCreated, "Button should have been created");
  assert.equal(buttonCreated.getAttribute("data-tfl-seen"), "false");
  assert.ok(buttonCreated.className.includes("unseen"));

  // Click the eye button to mark seen
  let eventPrevented = false;
  buttonCreated.dispatchEvent({
    type: "click",
    preventDefault: () => {
      eventPrevented = true;
    },
    stopPropagation: () => {},
  });

  assert.equal(eventPrevented, true);
  assert.equal(buttonCreated.getAttribute("data-tfl-seen"), "true");
  assert.ok(buttonCreated.className.includes("seen"));
});
