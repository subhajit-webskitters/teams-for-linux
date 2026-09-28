const { COMPOSE_SELECTORS, findCompose } = require("../../helpers/composeBox");
const { isTeamsHost } = require("../../helpers/teamsHosts");

const BYTE_PATTERN = /^[01]{8}$/u;
const MIN_BINARY_BYTES = 2;
const COMPOSER_SETTLE_MS = 100;
const UI_ATTRIBUTE = "data-tfl-binary-ui";

const SEND_BUTTON_SELECTORS = [
  'button[data-tid="sendMessageCommands-send"]',
  'button[data-tid="send-message"]',
  'button[data-tid="sendMessage"]',
  'button[data-tid*="send-message" i]',
  'button[data-testid*="send-message" i]',
  'button[aria-label^="Send (" i]',
  'button[aria-label="Send" i]',
  'button[title="Send" i]',
];

const MESSAGE_CONTAINER_SELECTORS = [
  '[data-tid="chat-pane-message"]',
  '[data-tid="channel-pane-message"]',
  '[data-tid="thread-pane-message"]',
];

const MESSAGE_BODY_SELECTORS = [
  '[data-tid="messageBodyContent"]',
  '[data-tid="message-body"]',
  '[data-tid*="message-body" i]',
  '[data-tid*="chat-message-text" i]',
  '[data-testid*="message-body" i]',
  '[id^="message-body-"]',
];

const QUOTED_CONTENT_SELECTORS = [
  "blockquote",
  '[data-tid*="quoted" i]',
  '[data-tid*="reply-preview" i]',
  '[data-tid*="forwarded" i]',
  '[data-testid*="quoted" i]',
  '[data-testid*="reply-preview" i]',
  '[aria-label*="quoted message" i]',
  '[aria-label*="reply preview" i]',
];

const UNSUPPORTED_COMPOSER_CONTENT_SELECTORS = [
  "img",
  "video",
  "audio",
  "object",
  "embed",
  '[data-tid*="mention" i]',
  '[data-tid*="attachment" i]',
  '[contenteditable="false"][role="button"]',
];

const SEND_BUTTON_SELECTOR = SEND_BUTTON_SELECTORS.join(",");
const MESSAGE_CONTAINER_SELECTOR = MESSAGE_CONTAINER_SELECTORS.join(",");
const MESSAGE_BODY_SELECTOR = MESSAGE_BODY_SELECTORS.join(",");
const QUOTED_CONTENT_SELECTOR = QUOTED_CONTENT_SELECTORS.join(",");
const UNSUPPORTED_COMPOSER_CONTENT_SELECTOR =
  UNSUPPORTED_COMPOSER_CONTENT_SELECTORS.join(",");

/**
 * Normalize whitespace, non-breaking spaces, and line breaks to single spaces.
 *
 * @param {string} text
 * @returns {string}
 */
function normalizeWhitespace(text) {
  if (typeof text !== "string") {
    return "";
  }
  return text.replace(/[\u00A0\s]+/gu, " ").trim();
}

/**
 * Convert text to space-separated UTF-8 bytes represented as 8-bit binary.
 * Binary messaging is encoding only; it does not encrypt or protect content.
 *
 * @param {string} text
 * @returns {string}
 */
function encodeTextToBinary(text) {
  if (typeof text !== "string") {
    throw new TypeError("Text to encode must be a string");
  }

  return Array.from(new TextEncoder().encode(text), (byte) =>
    byte.toString(2).padStart(8, "0")
  ).join(" ");
}

function parseBinaryBytes(binary) {
  if (typeof binary !== "string") {
    throw new TypeError("Binary text to decode must be a string");
  }

  const trimmed = binary.trim();
  if (trimmed === "") {
    return [];
  }

  const groups = trimmed.split(/\s+/u);
  if (!groups.every((group) => BYTE_PATTERN.test(group))) {
    throw new TypeError("Binary text must contain only 8-bit groups");
  }

  return groups.map((group) => Number.parseInt(group, 2));
}

/**
 * Decode space/newline-separated 8-bit groups as strict UTF-8.
 *
 * @param {string} binary
 * @returns {string}
 */
function decodeBinaryToText(binary) {
  const bytes = parseBinaryBytes(binary);
  return new TextDecoder("utf-8", { fatal: true }).decode(
    Uint8Array.from(bytes)
  );
}

/**
 * Conservative binary-message classification. Requiring at least two bytes,
 * strict UTF-8, a visible decoded character, and canonical round-tripping
 * avoids treating ordinary numeric text as an encoded message.
 *
 * @param {string} text
 * @returns {boolean}
 */
function isLikelyBinaryMessage(text) {
  try {
    const bytes = parseBinaryBytes(text);
    if (bytes.length < MIN_BINARY_BYTES) {
      return false;
    }

    const decoded = new TextDecoder("utf-8", { fatal: true }).decode(
      Uint8Array.from(bytes)
    );
    if (!/\S/u.test(decoded)) {
      return false;
    }
    // These are the exact non-printing byte values binary detection rejects.
    // eslint-disable-next-line no-control-regex
    if (/[\u0000-\u0008\u000B\u000C\u000E-\u001F\u007F]/u.test(decoded)) {
      return false;
    }

    const canonicalInput = bytes
      .map((byte) => byte.toString(2).padStart(8, "0"))
      .join(" ");
    return encodeTextToBinary(decoded) === canonicalInput;
  } catch {
    return false;
  }
}

function isElement(node) {
  return Boolean(node && node.nodeType === 1);
}

function collectMatches(root, selector) {
  if (!root) return [];

  const matches = [];
  if (isElement(root) && root.matches(selector)) {
    matches.push(root);
  }
  if (typeof root.querySelectorAll === "function") {
    matches.push(...root.querySelectorAll(selector));
  }
  return matches;
}

function isVisible(element) {
  if (!element || element.getAttribute("aria-hidden") === "true") {
    return false;
  }
  if (typeof element.getClientRects !== "function") {
    return true;
  }
  return element.getClientRects().length > 0;
}

function composerText(composer) {
  return typeof composer.innerText === "string"
    ? composer.innerText
    : composer.textContent || "";
}

class BinaryMessagingController {
  #document;
  #window;
  #MutationObserver;
  #InputEvent;
  #hostname;
  #observer = null;
  #started = false;
  #modeEnabled = false;
  #settings = null;
  #controls = new WeakMap();
  #messageStates = new WeakMap();
  #pendingEncoded = new WeakMap();
  #scheduledSends = new WeakSet();
  #replayedSendButtons = new WeakSet();
  #pendingRoots = new Set();
  #scanScheduled = false;
  #lastComposer = null;
  #internalEdit = false;

  constructor(environment = {}) {
    this.#document = environment.document ?? globalThis.document;
    this.#window = environment.window ?? globalThis;
    this.#MutationObserver =
      environment.MutationObserver ?? globalThis.MutationObserver;
    this.#InputEvent = environment.InputEvent ?? globalThis.InputEvent;
    this.#hostname =
      environment.hostname ?? this.#window?.location?.hostname ?? "";

    this.onClick = this.onClick.bind(this);
    this.onKeyDown = this.onKeyDown.bind(this);
    this.onFocusIn = this.onFocusIn.bind(this);
    this.onInput = this.onInput.bind(this);
    this.stop = this.stop.bind(this);
  }

  init(config = {}) {
    if (this.#started || !this.#document?.body || !isTeamsHost(this.#hostname)) {
      return;
    }

    const configured = config.binaryMessaging ?? {};
    this.#settings = {
      enabled: configured.enabled === true,
      autoDetectReceivedMessages:
        configured.autoDetectReceivedMessages !== false,
      showTranslateButton: configured.showTranslateButton !== false,
    };
    this.#modeEnabled = this.#settings.enabled;
    this.#started = true;

    this.installStyles();
    this.discover(this.#document.body);

    this.#document.addEventListener("click", this.onClick, true);
    this.#document.addEventListener("keydown", this.onKeyDown, true);
    this.#document.addEventListener("focusin", this.onFocusIn, true);
    this.#document.addEventListener("input", this.onInput, true);
    this.#window.addEventListener?.("pagehide", this.stop, { once: true });

    this.#observer = new this.#MutationObserver((records) => {
      for (const record of records) {
        if (record.type === "characterData") {
          const parent = record.target.parentElement;
          if (parent && !this.isBinaryUi(parent)) {
            this.queueDiscovery(parent);
          }
          continue;
        }

        if (isElement(record.target) && !this.isBinaryUi(record.target)) {
          const changedBody = this.findMessageBodyForNode(record.target);
          if (changedBody) {
            this.queueDiscovery(changedBody);
          }
        }
        for (const node of record.addedNodes) {
          if (isElement(node) && !this.isBinaryUi(node)) {
            this.queueDiscovery(node);
          }
        }
      }
    });
    this.#observer.observe(this.#document.body, {
      childList: true,
      characterData: true,
      subtree: true,
    });
  }

  stop() {
    if (!this.#started) return;
    this.#started = false;
    this.#observer?.disconnect();
    this.#observer = null;
    this.#pendingRoots.clear();
    this.#scanScheduled = false;
    this.#document.removeEventListener("click", this.onClick, true);
    this.#document.removeEventListener("keydown", this.onKeyDown, true);
    this.#document.removeEventListener("focusin", this.onFocusIn, true);
    this.#document.removeEventListener("input", this.onInput, true);
    this.#window.removeEventListener?.("pagehide", this.stop);
  }

  installStyles() {
    if (this.#document.getElementById("tfl-binary-messaging-styles")) return;

    const style = this.#document.createElement("style");
    style.id = "tfl-binary-messaging-styles";
    style.setAttribute(UI_ATTRIBUTE, "true");
    style.textContent = `
      .tfl-binary-mode-host {
        display: inline-flex;
        align-items: center;
        margin-inline: 4px;
      }
      .tfl-binary-mode-toggle,
      .tfl-binary-translate-button {
        box-sizing: border-box;
        border: 1px solid currentColor;
        border-radius: 4px;
        background: transparent;
        color: inherit;
        cursor: pointer;
        font: inherit;
        font-size: 12px;
        line-height: 1.25;
        padding: 4px 7px;
        opacity: .78;
      }
      .tfl-binary-mode-toggle:hover,
      .tfl-binary-mode-toggle:focus-visible,
      .tfl-binary-translate-button:hover,
      .tfl-binary-translate-button:focus-visible {
        opacity: 1;
        outline: 2px solid currentColor;
        outline-offset: 2px;
      }
      .tfl-binary-mode-toggle[aria-pressed="true"] {
        background: #5b5fc7;
        border-color: #5b5fc7;
        color: #fff;
        opacity: 1;
      }
      .tfl-binary-translation {
        align-items: flex-start;
        display: flex;
        flex-direction: column;
        gap: 4px;
        margin-block-start: 6px;
      }
      .tfl-binary-decoded {
        border-inline-start: 3px solid #5b5fc7;
        padding-inline-start: 8px;
        white-space: pre-wrap;
      }
      .tfl-binary-decoded-label {
        font-weight: 600;
      }
    `;
    (this.#document.head ?? this.#document.body).append(style);
  }

  queueDiscovery(root) {
    this.#pendingRoots.add(root);
    if (this.#scanScheduled) return;
    this.#scanScheduled = true;

    const schedule = this.#window.queueMicrotask ?? queueMicrotask;
    schedule(() => {
      if (!this.#started) {
        this.#pendingRoots.clear();
        this.#scanScheduled = false;
        return;
      }
      this.#scanScheduled = false;
      const roots = Array.from(this.#pendingRoots);
      this.#pendingRoots.clear();
      for (const pendingRoot of roots) {
        this.discover(pendingRoot);
      }
    });
  }

  discover(root) {
    if (!root || this.isBinaryUi(root)) return;

    const composeSelector = COMPOSE_SELECTORS.join(",");
    for (const composer of collectMatches(root, composeSelector)) {
      if (this.isUsableComposer(composer)) {
        this.ensureModeControl(composer);
      }
    }
    for (const sendButton of collectMatches(root, SEND_BUTTON_SELECTOR)) {
      const composer = this.findComposerForSend(sendButton);
      if (composer) {
        this.ensureModeControl(composer);
      }
    }

    if (
      !this.#settings.autoDetectReceivedMessages ||
      !this.#settings.showTranslateButton
    ) {
      return;
    }

    const rawBodies = [];
    const closestBody = this.findMessageBodyForNode(root);
    if (closestBody) {
      rawBodies.push(closestBody);
    }
    for (const body of collectMatches(root, MESSAGE_BODY_SELECTOR)) {
      rawBodies.push(body);
    }
    for (const container of collectMatches(root, MESSAGE_CONTAINER_SELECTOR)) {
      const body = this.findMessageBodyInContainer(container);
      if (body) {
        rawBodies.push(body);
      }
    }

    const messageBodies = new Set();
    for (const raw of rawBodies) {
      const innermost = this.findInnermostBody(raw);
      if (innermost) {
        messageBodies.add(innermost);
      }
    }

    for (const body of messageBodies) {
      if (body.querySelector(MESSAGE_BODY_SELECTOR)) {
        continue;
      }
      this.processMessageBody(body);
    }
  }

  findInnermostBody(body) {
    if (!isElement(body)) return null;
    let current = body;
    let deeper = current.querySelector(MESSAGE_BODY_SELECTOR);
    while (deeper && deeper !== current) {
      current = deeper;
      deeper = current.querySelector(MESSAGE_BODY_SELECTOR);
    }
    return current;
  }

  findMessageBodyInContainer(container) {
    if (!isElement(container)) return null;

    const knownBody = container.querySelector(MESSAGE_BODY_SELECTOR);
    if (knownBody) return this.findInnermostBody(knownBody);

    return (
      Array.from(container.children).find(
        (child) =>
          child.matches('div[dir="auto"]') &&
          !child.matches(QUOTED_CONTENT_SELECTOR)
      ) ?? null
    );
  }

  findMessageBodyForNode(node) {
    if (!isElement(node)) return null;

    const knownBody = node.closest(MESSAGE_BODY_SELECTOR);
    if (knownBody) return this.findInnermostBody(knownBody);

    const container = node.closest(MESSAGE_CONTAINER_SELECTOR);
    return this.findMessageBodyInContainer(container);
  }

  isBinaryUi(node) {
    return Boolean(
      isElement(node) &&
        (node.hasAttribute(UI_ATTRIBUTE) || node.closest(`[${UI_ATTRIBUTE}]`))
    );
  }

  isUsableComposer(composer) {
    return (
      isElement(composer) &&
      !this.isBinaryUi(composer) &&
      !composer.closest(MESSAGE_BODY_SELECTOR) &&
      !composer.closest(MESSAGE_CONTAINER_SELECTOR) &&
      isVisible(composer)
    );
  }

  findSendButton(composer) {
    let current = composer.parentElement;
    for (let depth = 0; current && depth < 10; depth += 1) {
      const button = current.querySelector(SEND_BUTTON_SELECTOR);
      if (button && isVisible(button)) {
        return button;
      }
      current = current.parentElement;
    }
    return null;
  }

  findComposerForSend(sendButton) {
    let current = sendButton.parentElement;
    for (let depth = 0; current && depth < 10; depth += 1) {
      if (
        this.#lastComposer &&
        current.contains(this.#lastComposer) &&
        this.isUsableComposer(this.#lastComposer)
      ) {
        return this.#lastComposer;
      }
      const composer = findCompose(current, COMPOSE_SELECTORS);
      if (this.isUsableComposer(composer)) {
        return composer;
      }
      current = current.parentElement;
    }
    return null;
  }

  closestComposer(target) {
    if (!isElement(target)) return null;
    const composer = target.closest(COMPOSE_SELECTORS.join(","));
    return this.isUsableComposer(composer) ? composer : null;
  }

  ensureModeControl(composer) {
    const existing = this.#controls.get(composer);
    if (existing?.isConnected) return;

    const sendButton = this.findSendButton(composer);
    if (!sendButton?.parentElement) return;

    const host = this.#document.createElement("span");
    host.className = "tfl-binary-mode-host";
    host.setAttribute(UI_ATTRIBUTE, "true");

    const button = this.#document.createElement("button");
    button.type = "button";
    button.className = "tfl-binary-mode-toggle";
    button.setAttribute(UI_ATTRIBUTE, "true");
    button.addEventListener("click", (event) => {
      event.preventDefault();
      event.stopPropagation();
      this.#modeEnabled = !this.#modeEnabled;
      this.updateModeControls();
    });
    host.append(button);
    sendButton.parentElement.insertBefore(host, sendButton);
    this.#controls.set(composer, host);
    this.updateModeControls();
  }

  updateModeControls() {
    const buttons = this.#document.querySelectorAll(
      `.tfl-binary-mode-toggle[${UI_ATTRIBUTE}]`
    );
    for (const button of buttons) {
      const state = this.#modeEnabled ? "ON" : "OFF";
      button.textContent = `Binary Mode ${state}`;
      button.title = `Binary Messaging Mode is ${state}. Activate to turn it ${
        this.#modeEnabled ? "off" : "on"
      }.`;
      button.setAttribute(
        "aria-label",
        `Binary Messaging Mode ${state}. Activate to turn ${
          this.#modeEnabled ? "off" : "on"
        }.`
      );
      button.setAttribute("aria-pressed", String(this.#modeEnabled));
    }
  }

  onFocusIn(event) {
    const composer = this.closestComposer(event.target);
    if (composer) {
      this.#lastComposer = composer;
    }
  }

  onInput(event) {
    if (this.#internalEdit) return;
    const composer = this.closestComposer(event.target);
    if (composer) {
      this.#pendingEncoded.delete(composer);
    }
  }

  onClick(event) {
    if (!isElement(event.target)) return;

    const sendButton = event.target.closest(SEND_BUTTON_SELECTOR);
    if (sendButton && this.#replayedSendButtons.has(sendButton)) {
      this.#replayedSendButtons.delete(sendButton);
      return;
    }
    if (!this.#modeEnabled) return;

    if (
      !sendButton ||
      sendButton.disabled ||
      sendButton.getAttribute("aria-disabled") === "true"
    ) {
      return;
    }

    const composer = this.findComposerForSend(sendButton);
    if (composer && this.transformComposer(composer)) {
      event.preventDefault();
      event.stopImmediatePropagation();
      this.scheduleSend(composer, sendButton);
    }
  }

  onKeyDown(event) {
    if (
      !this.#modeEnabled ||
      event.key !== "Enter" ||
      event.shiftKey ||
      event.altKey ||
      event.isComposing
    ) {
      return;
    }

    const composer = this.closestComposer(event.target);
    if (!composer) return;

    const sendButton = this.findSendButton(composer);
    if (
      !sendButton ||
      sendButton.disabled ||
      sendButton.getAttribute("aria-disabled") === "true"
    ) {
      return;
    }

    const shortcutLabel = [
      sendButton.getAttribute("aria-label"),
      sendButton.getAttribute("title"),
    ]
      .filter(Boolean)
      .join(" ");
    const expectsControl = /(?:ctrl|control)\s*\+\s*enter/iu.test(
      shortcutLabel
    );
    const expectsMeta = /(?:cmd|command|⌘)\s*\+\s*enter/iu.test(
      shortcutLabel
    );
    const isModifierEnter = event.ctrlKey || event.metaKey;
    if (!isModifierEnter && (expectsControl || expectsMeta)) {
      return;
    }
    if (this.transformComposer(composer)) {
      event.preventDefault();
      event.stopImmediatePropagation();
      this.scheduleSend(composer, sendButton);
    }
  }

  scheduleSend(composer, sendButton) {
    if (this.#scheduledSends.has(composer)) return;

    const encoded = this.#pendingEncoded.get(composer);
    if (!encoded) return;
    this.#scheduledSends.add(composer);

    const tryClickSend = (retriesLeft = 5) => {
      if (!composer.isConnected || !sendButton.isConnected) {
        this.#scheduledSends.delete(composer);
        return;
      }

      const current = composerText(composer);
      const isBinary =
        normalizeWhitespace(current) === normalizeWhitespace(encoded) ||
        isLikelyBinaryMessage(current);

      if (!isBinary) {
        this.#scheduledSends.delete(composer);
        return;
      }

      const isDisabled =
        sendButton.disabled ||
        sendButton.getAttribute("aria-disabled") === "true";

      if (isDisabled && retriesLeft > 0) {
        this.#window.setTimeout?.(() => tryClickSend(retriesLeft - 1), 50);
        return;
      }

      try {
        this.#replayedSendButtons.add(sendButton);
        sendButton.click();
      } finally {
        this.#replayedSendButtons.delete(sendButton);
        this.#scheduledSends.delete(composer);
      }
    };

    this.#window.setTimeout?.(() => tryClickSend(), COMPOSER_SETTLE_MS);
  }

  transformComposer(composer) {
    const text = composerText(composer);
    if (text === "") return false;

    if (composer.querySelector(UNSUPPORTED_COMPOSER_CONTENT_SELECTOR)) {
      return false;
    }

    const normText = normalizeWhitespace(text);
    const pending = this.#pendingEncoded.get(composer);
    if (
      pending &&
      (text === pending || normText === normalizeWhitespace(pending))
    ) {
      return true;
    }

    // Do NOT double-encode if the text is ALREADY a valid binary message!
    if (isLikelyBinaryMessage(text)) {
      this.#pendingEncoded.set(composer, text);
      return true;
    }

    const binary = encodeTextToBinary(text);
    if (!this.replaceComposerText(composer, binary)) {
      return false;
    }
    this.#pendingEncoded.set(composer, binary);
    this.#window.setTimeout?.(() => {
      if (this.#pendingEncoded.get(composer) === binary) {
        this.#pendingEncoded.delete(composer);
      }
    }, 15000);
    return true;
  }

  replaceComposerText(composer, text) {
    const editor = composer.ckeditorInstance;
    if (
      editor?.model?.document?.getRoot &&
      typeof editor.model.change === "function"
    ) {
      return this.replaceCkEditorText(composer, editor, text);
    }

    const selection = this.#window.getSelection?.();
    if (!selection || typeof this.#document.createRange !== "function") {
      return false;
    }

    composer.focus();
    const range = this.#document.createRange();
    range.selectNodeContents(composer);
    selection.removeAllRanges();
    selection.addRange(range);

    this.#internalEdit = true;
    try {
      let replaced = false;
      if (typeof this.#document.execCommand === "function") {
        replaced = this.#document.execCommand("insertText", false, text);
      }

      if (!replaced) {
        const beforeInput = new this.#InputEvent("beforeinput", {
          bubbles: true,
          cancelable: true,
          composed: true,
          data: text,
          inputType: "insertText",
        });
        if (!composer.dispatchEvent(beforeInput)) {
          return false;
        }

        range.deleteContents();
        const textNode = this.#document.createTextNode(text);
        range.insertNode(textNode);
        range.setStartAfter(textNode);
        range.collapse(true);
        selection.removeAllRanges();
        selection.addRange(range);
      }

      composer.dispatchEvent(
        new this.#InputEvent("input", {
          bubbles: true,
          composed: true,
          data: text,
          inputType: "insertText",
        })
      );
      const current = composerText(composer);
      return (
        normalizeWhitespace(current) === normalizeWhitespace(text) ||
        isLikelyBinaryMessage(current)
      );
    } finally {
      this.#internalEdit = false;
    }
  }

  replaceCkEditorText(composer, editor, text) {
    this.#internalEdit = true;
    try {
      const selectAllCommand = editor.commands?.get?.("selectAll");
      const insertTextCommand = editor.commands?.get?.("insertText");

      if (
        typeof editor.execute === "function" &&
        selectAllCommand?.isEnabled !== false &&
        insertTextCommand?.isEnabled !== false &&
        selectAllCommand &&
        insertTextCommand
      ) {
        // Teams commits outgoing messages from CKEditor's command pipeline,
        // not directly from the rendered DOM. Use the same command path as
        // native editing so Teams' compose service receives the binary draft.
        editor.execute("selectAll");
        editor.execute("insertText", { text });
      } else {
        editor.model.change((writer) => {
          const root = editor.model.document.getRoot();
          writer.remove(writer.createRangeIn(root));
          const paragraph = writer.createElement("paragraph");
          writer.append(paragraph, root);
          writer.insertText(text, paragraph, 0);
          writer.setSelection(paragraph, "end");
        });
      }
      editor.editing?.view?.focus?.();

      // Programmatic CKEditor updates change the editable DOM, but Teams
      // also keeps composer state outside the editor. Notify that integration
      // before replaying Send so it cannot submit the stale, readable draft.
      composer.dispatchEvent(
        new this.#InputEvent("input", {
          bubbles: true,
          composed: true,
          data: text,
          inputType: "insertReplacementText",
        })
      );
      const current = composerText(composer);
      return (
        normalizeWhitespace(current) === normalizeWhitespace(text) ||
        isLikelyBinaryMessage(current) ||
        Boolean(editor.model)
      );
    } catch {
      return false;
    } finally {
      this.#internalEdit = false;
    }
  }

  messageText(body) {
    const clone = body.cloneNode(true);
    for (const node of clone.querySelectorAll(
      `${QUOTED_CONTENT_SELECTOR},[${UI_ATTRIBUTE}]`
    )) {
      node.remove();
    }
    return (clone.textContent || "").trim();
  }

  findDeepestBinaryElement(element) {
    if (!isElement(element)) return null;
    let current = element;
    let deeper = true;
    while (deeper) {
      deeper = false;
      if (current.children) {
        for (const child of current.children) {
          if (!isElement(child) || child.hasAttribute?.(UI_ATTRIBUTE)) continue;
          const text = this.messageText(child);
          if (isLikelyBinaryMessage(text)) {
            current = child;
            deeper = true;
            break;
          }
        }
      }
    }
    return current;
  }

  processMessageBody(body) {
    if (
      !isElement(body) ||
      this.isBinaryUi(body) ||
      body.closest(COMPOSE_SELECTORS.join(",")) ||
      body.closest(QUOTED_CONTENT_SELECTOR)
    ) {
      return;
    }

    // Always resolve to the deepest/innermost element that contains the binary text
    const leafBody = this.findDeepestBinaryElement(body);
    if (!leafBody) return;
    if (leafBody !== body) {
      // body is an outer container or bubble — process only the leaf inside it
      this.processMessageBody(leafBody);
      return;
    }

    const binary = this.messageText(body);
    const previous = this.#messageStates.get(body);
    if (previous?.binary === binary && previous.host?.isConnected) {
      return;
    }

    if (!isLikelyBinaryMessage(binary)) {
      previous?.host?.remove();
      this.#messageStates.delete(body);
      return;
    }

    const decoded = decodeBinaryToText(binary);
    let state = previous;
    if (!state?.host?.isConnected) {
      state = this.createTranslationAction(body);
      if (!state) return;
    }

    // Deduplicate: remove any stray or outer translation actions in the ancestor tree
    let scope = body.parentElement;
    for (let depth = 0; scope && depth < 8; depth += 1) {
      const translations = scope.querySelectorAll(
        `.tfl-binary-translation[${UI_ATTRIBUTE}]`
      );
      for (const el of translations) {
        if (el !== state.host) {
          el.remove();
        }
      }
      if (
        scope.matches?.(MESSAGE_CONTAINER_SELECTOR) ||
        scope.matches?.('[role="listitem"]') ||
        scope.matches?.('[data-tid*="message"]')
      ) {
        break;
      }
      scope = scope.parentElement;
    }

    state.binary = binary;
    state.decodedText.textContent = decoded;
    this.#messageStates.set(body, state);
  }

  createTranslationAction(body) {
    if (!body.parentElement) return null;

    const host = this.#document.createElement("div");
    host.className = "tfl-binary-translation";
    host.setAttribute(UI_ATTRIBUTE, "true");

    const button = this.#document.createElement("button");
    button.type = "button";
    button.className = "tfl-binary-translate-button";
    button.textContent = "Translate Binary";
    button.title = "Decode this UTF-8 binary message locally";
    button.setAttribute("aria-expanded", "false");
    button.setAttribute(UI_ATTRIBUTE, "true");

    const decoded = this.#document.createElement("div");
    decoded.className = "tfl-binary-decoded";
    decoded.hidden = true;
    decoded.setAttribute("role", "status");
    decoded.setAttribute("aria-live", "polite");
    decoded.setAttribute(UI_ATTRIBUTE, "true");

    const label = this.#document.createElement("span");
    label.className = "tfl-binary-decoded-label";
    label.textContent = "Decoded: ";
    label.setAttribute(UI_ATTRIBUTE, "true");
    const decodedText = this.#document.createElement("span");
    decodedText.setAttribute(UI_ATTRIBUTE, "true");
    decoded.append(label, decodedText);

    button.addEventListener("click", (event) => {
      event.preventDefault();
      event.stopPropagation();
      decoded.hidden = !decoded.hidden;
      const expanded = !decoded.hidden;
      button.textContent = expanded ? "Hide Translation" : "Translate Binary";
      button.setAttribute("aria-expanded", String(expanded));
    });

    host.append(button, decoded);
    body.parentElement.insertBefore(host, body.nextSibling);
    return { binary: "", host, button, decoded, decodedText };
  }
}

const controller = new BinaryMessagingController();

module.exports = {
  init: (config) => controller.init(config),
  stop: () => controller.stop(),
  BinaryMessagingController,
  encodeTextToBinary,
  decodeBinaryToText,
  isLikelyBinaryMessage,
  normalizeWhitespace,
};
