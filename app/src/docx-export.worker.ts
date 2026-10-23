import type { JSONContent } from "@tiptap/core";
import { exportDocx } from "./docx";

self.onmessage = async (event: MessageEvent<JSONContent>) => {
  try {
    const blob = await exportDocx(event.data);
    const bytes = await blob.arrayBuffer();
    self.postMessage({ ok: true, bytes }, { transfer: [bytes] });
  } catch {
    self.postMessage({ ok: false });
  }
};
