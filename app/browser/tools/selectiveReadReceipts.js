"use strict";

const { isTeamsHost } = require("../../helpers/teamsHosts");

const UI_ATTRIBUTE = "data-tfl-receipt-ui";

const MESSAGE_SELECTORS = [
  '[data-tid="chat-pane-message"]',
  '[data-tid="channel-pane-message"]',
  '[data-tid="thread-pane-message"]',
  ".fui-ChatMessage",
  '[class*="ChatMessage"]',
  '[class*="ui-chat__item__message"]',
];

const TIME_SELECTORS = [
  "time",
  '[data-tid*="timestamp" i]',
  '[data-tid*="message-time" i]',
  '[class*="timestamp" i]',
  '[class*="time" i]',
];

const MY_MESSAGE_SELECTORS = [
  ".fui-ChatMyMessage",
  '[class*="ChatMyMessage"]',
  '[class*="--mine"]',
  '[data-tid*="mymessage" i]',
  '[data-tid*="my-message" i]',
  '[class*="ui-chat__item__message--mine"]',
];

const EYE_OFF_SVG = `<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M17.94 17.94A10.07 10.07 0 0 1 12 20c-7 0-11-8-11-8a18.45 18.45 0 0 1 5.06-5.94M9.9 4.24A9.12 9.12 0 0 1 12 4c7 0 11 8 11 8a18.5 18.5 0 0 1-2.16 3.19m-6.72-1.07a3 3 0 1 1-4.24-4.24"></path><line x1="1" y1="1" x2="23" y2="23"></line></svg>`;

const EYE_ON_SVG = `<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"></path><circle cx="12" cy="12" r="3"></circle></svg>`;

/**
 * Check whether a given URL is a Teams read receipt / consumption horizon endpoint.
 *
 * @param {string} url
 * @returns {boolean}
 */
function isReadReceiptUrl(url) {
  if (typeof url !== "string") return false;
  const lower = url.toLowerCase();
  return (
    lower.includes("/consumptionhorizon") ||
    lower.includes("/consumptionhorizons") ||
    lower.includes("/readreceipt") ||
    lower.includes("/readreceipts")
  );
}

/**
 * Extract conversation ID from a request URL or URI string.
 *
 * @param {string} url
 * @returns {string}
 */
function extractConversationId(url) {
  if (typeof url !== "string") return "default";
  const threadMatch = url.match(/(19:[a-zA-Z0-9_:.@%-]+)/i);
  if (threadMatch) {
    return decodeURIComponent(threadMatch[1]);
  }
  const match = url.match(/(?:conversations|chats?)[/:]([a-zA-Z0-9_:.@%-]+)/i);
  if (match) {
    return decodeURIComponent(match[1]);
  }
  return "default";
}

function getMessageId(element) {
  if (!isElement(element)) return null;
  return (
    element.getAttribute("data-mid") ||
    element.getAttribute("data-message-id") ||
    element.id ||
    null
  );
}

function isElement(node) {
  return Boolean(node && node.nodeType === 1);
}

class SelectiveReadReceiptsController {
  #document;
  #window;
  #MutationObserver;
  #Response;
  #hostname;
  #observer = null;
  #started = false;
  #settings = null;
  #pendingReceipts = new Map();
  #autoSendPending = new Set();
  #seenMessages = new WeakSet();
  #seenMessageIds = new Set();
  #interceptorInstalled = false;
  #originalFetch = null;
  #originalXhrOpen = null;
  #originalXhrSend = null;

  constructor(environment = {}) {
    this.#document = environment.document ?? globalThis.document;
    this.#window = environment.window ?? globalThis;
    this.#MutationObserver =
      environment.MutationObserver ?? globalThis.MutationObserver;
    this.#Response = environment.Response ?? globalThis.Response;
    this.#hostname =
      environment.hostname ?? this.#window?.location?.hostname ?? "";

    this.stop = this.stop.bind(this);
  }

  get pendingCount() {
    return this.#pendingReceipts.size;
  }

  get isStarted() {
    return this.#started;
  }

  isMyMessage(element) {
    if (!isElement(element)) return false;
    for (const selector of MY_MESSAGE_SELECTORS) {
      if (element.matches?.(selector) || element.closest?.(selector)) {
        return true;
      }
    }
    return false;
  }

  installStyles() {
    if (this.#document.getElementById("tfl-selective-read-receipt-styles")) return;

    const style = this.#document.createElement("style");
    style.id = "tfl-selective-read-receipt-styles";
    style.setAttribute(UI_ATTRIBUTE, "true");
    style.textContent = `
      .tfl-read-receipt-btn {
        display: inline-flex;
        align-items: center;
        justify-content: center;
        background: transparent;
        border: 1px solid transparent;
        cursor: pointer;
        padding: 2px 4px;
        border-radius: 4px;
        vertical-align: middle;
        margin-inline: 4px;
        line-height: 1;
        transition: color 0.15s ease, opacity 0.15s ease, background 0.15s ease;
      }
      .tfl-read-receipt-btn.unseen {
        color: #a19f9d;
        opacity: 0.65;
      }
      .tfl-read-receipt-btn.unseen:hover {
        opacity: 1;
        color: #e1dfdd;
        background: rgba(128, 128, 128, 0.2);
        border-color: rgba(128, 128, 128, 0.3);
      }
      .tfl-read-receipt-btn.seen {
        color: #7f85f5;
        opacity: 1;
      }
      .tfl-read-receipt-btn.seen:hover {
        background: rgba(127, 133, 245, 0.2);
      }
    `;
    (this.#document.head ?? this.#document.body).append(style);
  }

  installNetworkInterceptor() {
    if (this.#interceptorInstalled || !this.#window?.fetch) return;
    this.#interceptorInstalled = true;

    const originalFetch = this.#window.fetch;
    this.#originalFetch = originalFetch;
    const controller = this;

    this.#window.fetch = async function (input, init) {
      let url = "";
      if (typeof input === "string") {
        url = input;
      } else if (input && typeof input.url === "string") {
        url = input.url;
      } else if (input && typeof input.href === "string") {
        url = input.href;
      } else if (input && typeof input.toString === "function") {
        url = input.toString();
      }

      if (isReadReceiptUrl(url)) {
        const conversationId = extractConversationId(url);

        // If the user already clicked "Seen" for this conversation, release immediately
        if (
          controller.#autoSendPending.has(conversationId) ||
          controller.#autoSendPending.has("default")
        ) {
          controller.#autoSendPending.delete(conversationId);
          controller.#autoSendPending.delete("default");
          return originalFetch.apply(this, arguments);
        }

        let bufferedInput = input;
        if (
          typeof input === "object" &&
          input !== null &&
          typeof input.clone === "function"
        ) {
          try {
            bufferedInput = input.clone();
          } catch {
            // Keep input if cloning fails
          }
        }

        // Buffer the request so user can manually send it
        controller.#pendingReceipts.set(conversationId, {
          input: bufferedInput,
          init: init ? { ...init } : undefined,
          timestamp: Date.now(),
        });

        // Return a mock 200 OK so Teams does not log a network failure
        const ResponseConstructor = controller.#Response ?? globalThis.Response;
        if (typeof ResponseConstructor === "function") {
          return new ResponseConstructor(
            JSON.stringify({ success: true, consumptionhorizon: "buffered" }),
            {
              status: 200,
              statusText: "OK",
              headers: { "Content-Type": "application/json" },
            }
          );
        }

        return {
          ok: true,
          status: 200,
          json: async () => ({ success: true }),
          text: async () => '{"success":true}',
        };
      }

      return originalFetch.apply(this, arguments);
    };

    if (this.#window.XMLHttpRequest) {
      const originalOpen = this.#window.XMLHttpRequest.prototype.open;
      const originalSend = this.#window.XMLHttpRequest.prototype.send;
      this.#originalXhrOpen = originalOpen;
      this.#originalXhrSend = originalSend;

      this.#window.XMLHttpRequest.prototype.open = function (method, url) {
        this._tflUrl = url;
        this._tflMethod = method;
        return originalOpen.apply(this, arguments);
      };

      this.#window.XMLHttpRequest.prototype.send = function (body) {
        if (this._tflBypassInterceptor) {
          return originalSend.apply(this, arguments);
        }

        if (this._tflUrl && isReadReceiptUrl(this._tflUrl)) {
          const conversationId = extractConversationId(this._tflUrl);

          if (
            controller.#autoSendPending.has(conversationId) ||
            controller.#autoSendPending.has("default")
          ) {
            controller.#autoSendPending.delete(conversationId);
            controller.#autoSendPending.delete("default");
            return originalSend.apply(this, arguments);
          }

          controller.#pendingReceipts.set(conversationId, {
            xhrData: { url: this._tflUrl, method: this._tflMethod, body },
            timestamp: Date.now(),
          });

          // Simulate immediate 200 OK
          try {
            Object.defineProperty(this, "status", { value: 200, writable: false });
            Object.defineProperty(this, "readyState", { value: 4, writable: false });
            Object.defineProperty(this, "responseText", {
              value: '{"success":true}',
              writable: false,
            });
            setTimeout(() => {
              this.dispatchEvent(new Event("readystatechange"));
              this.dispatchEvent(new Event("load"));
              this.dispatchEvent(new Event("loadend"));
            }, 0);
          } catch {
            // Best effort synthetic response
          }
          return;
        }

        return originalSend.apply(this, arguments);
      };
    }
  }

  sendReadReceipt(conversationId = "default") {
    let receipt = this.#pendingReceipts.get(conversationId);
    if (!receipt && this.#pendingReceipts.size > 0) {
      // Fall back to most recent pending receipt
      const entries = Array.from(this.#pendingReceipts.entries());
      const lastEntry = entries[entries.length - 1];
      conversationId = lastEntry[0];
      receipt = lastEntry[1];
    }

    if (!receipt) {
      // Allow next receipt for this conversation through automatically
      this.#autoSendPending.add(conversationId);
      this.#autoSendPending.add("default");
      return false;
    }

    this.#pendingReceipts.delete(conversationId);

    try {
      if (receipt.input && this.#originalFetch) {
        this.#originalFetch.call(this.#window, receipt.input, receipt.init).catch((err) => {
          console.debug("[SELECTIVE_READ_RECEIPTS] Delivery failed:", err?.message);
        });
        return true;
      }
      if (receipt.xhrData && this.#window?.XMLHttpRequest) {
        const xhr = new this.#window.XMLHttpRequest();
        xhr._tflBypassInterceptor = true;
        xhr.open(receipt.xhrData.method || "PUT", receipt.xhrData.url);
        xhr.send(receipt.xhrData.body);
        return true;
      }
    } catch (err) {
      console.debug("[SELECTIVE_READ_RECEIPTS] Error releasing receipt:", err?.message);
    }
    return false;
  }

  init(config = {}) {
    const currentHostname = this.#window?.location?.hostname ?? this.#hostname;
    if (this.#started || !this.#document?.body || !isTeamsHost(currentHostname)) {
      return;
    }

    const configured = config.selectiveReadReceipts ?? {};
    if (configured.enabled === false) {
      return;
    }

    this.#settings = {
      enabled: configured.enabled !== false,
    };
    this.#started = true;

    this.installStyles();
    this.installNetworkInterceptor();
    this.discover(this.#document.body);

    this.#observer = new this.#MutationObserver((records) => {
      for (const record of records) {
        for (const node of record.addedNodes) {
          if (isElement(node) && !node.hasAttribute?.(UI_ATTRIBUTE)) {
            this.discover(node);
          }
        }
      }
    });

    this.#observer.observe(this.#document.body, {
      childList: true,
      subtree: true,
    });
  }

  stop() {
    if (!this.#started) return;
    this.#started = false;
    this.#observer?.disconnect();
    this.#observer = null;
    this.#pendingReceipts.clear();
    this.#autoSendPending.clear();
    this.#seenMessageIds.clear();
  }

  discover(root) {
    if (!isElement(root) || root.hasAttribute?.(UI_ATTRIBUTE)) return;

    const messageSelector = MESSAGE_SELECTORS.join(",");
    const messages = [];

    if (root.matches?.(messageSelector)) {
      messages.push(root);
    }
    if (typeof root.querySelectorAll === "function") {
      messages.push(...root.querySelectorAll(messageSelector));
    }

    for (const msg of messages) {
      if (this.isMyMessage(msg)) {
        continue;
      }
      this.ensureEyeControl(msg);
    }
  }

  markPrecedingMessagesAsSeen(messageElement) {
    try {
      let prev = messageElement.previousElementSibling;
      while (prev) {
        const prevBtn = prev.querySelector?.(
          `button[${UI_ATTRIBUTE}][data-tfl-seen="false"]`
        );
        if (prevBtn) {
          const msgId = getMessageId(prev);
          this.#seenMessages.add(prev);
          if (msgId) {
            this.#seenMessageIds.add(msgId);
          }
          prevBtn.className = "tfl-read-receipt-btn seen";
          prevBtn.setAttribute("data-tfl-seen", "true");
          prevBtn.setAttribute("aria-label", "Message marked as seen");
          prevBtn.title = "Seen: Read receipt sent to sender.";
          prevBtn.innerHTML = EYE_ON_SVG;
        }
        prev = prev.previousElementSibling;
      }
    } catch {
      // Best-effort DOM update
    }
  }

  ensureEyeControl(messageElement) {
    if (!isElement(messageElement) || messageElement.querySelector(`[${UI_ATTRIBUTE}]`)) {
      return;
    }

    // Find the best insertion target (timestamp or header)
    const timeSelector = TIME_SELECTORS.join(",");
    let target = messageElement.querySelector(timeSelector);

    if (!target) {
      target = messageElement.querySelector(
        '[data-tid*="author" i], [class*="author" i], [class*="header" i]'
      );
    }

    const msgId = getMessageId(messageElement);
    const isAlreadySeen =
      (msgId && this.#seenMessageIds.has(msgId)) ||
      this.#seenMessages.has(messageElement);

    const button = this.#document.createElement("button");
    button.type = "button";
    button.setAttribute(UI_ATTRIBUTE, "true");

    if (isAlreadySeen) {
      button.className = "tfl-read-receipt-btn seen";
      button.setAttribute("data-tfl-seen", "true");
      button.setAttribute("aria-label", "Message marked as seen");
      button.title = "Seen: Read receipt sent to sender.";
      button.innerHTML = EYE_ON_SVG;
    } else {
      button.className = "tfl-read-receipt-btn unseen";
      button.setAttribute("data-tfl-seen", "false");
      button.setAttribute("aria-label", "Mark message as seen");
      button.title =
        "Unseen: Sender cannot see you read this. Click to send read receipt.";
      button.innerHTML = EYE_OFF_SVG;
    }

    button.addEventListener("click", (event) => {
      event.preventDefault();
      event.stopPropagation();

      const currentMsgId = getMessageId(messageElement);
      if (
        this.#seenMessages.has(messageElement) ||
        (currentMsgId && this.#seenMessageIds.has(currentMsgId))
      ) {
        return;
      }

      this.#seenMessages.add(messageElement);
      if (currentMsgId) {
        this.#seenMessageIds.add(currentMsgId);
      }
      button.className = "tfl-read-receipt-btn seen";
      button.setAttribute("data-tfl-seen", "true");
      button.setAttribute("aria-label", "Message marked as seen");
      button.title = "Seen: Read receipt sent to sender.";
      button.innerHTML = EYE_ON_SVG;

      this.markPrecedingMessagesAsSeen(messageElement);

      const convContainer = messageElement.closest?.(
        '[data-thread-id], [data-conversation-id], [id*="chat-message-list-"]'
      );
      const conversationId =
        convContainer?.getAttribute?.("data-thread-id") ||
        convContainer?.getAttribute?.("data-conversation-id") ||
        extractConversationId(this.#window?.location?.href || "");

      this.sendReadReceipt(conversationId);
    });

    if (target?.parentElement) {
      target.parentElement.insertBefore(button, target.nextSibling);
    } else {
      messageElement.prepend(button);
    }
  }
}

const controller = new SelectiveReadReceiptsController();

module.exports = {
  init: (config) => controller.init(config),
  stop: () => controller.stop(),
  SelectiveReadReceiptsController,
  isReadReceiptUrl,
  extractConversationId,
  getMessageId,
};
