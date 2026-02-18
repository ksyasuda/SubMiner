import type {
  KikuDuplicateCardInfo,
  KikuFieldGroupingChoice,
  KikuMergePreviewResponse,
} from "../../types";
import type { ModalStateReader, RendererContext } from "../context";

export function createKikuModal(
  ctx: RendererContext,
  options: {
    modalStateReader: Pick<ModalStateReader, "isAnyModalOpen">;
    syncSettingsModalSubtitleSuppression: () => void;
  },
) {
  function formatMediaMeta(card: KikuDuplicateCardInfo): string {
    const parts: string[] = [];
    parts.push(card.hasAudio ? "Audio: Yes" : "Audio: No");
    parts.push(card.hasImage ? "Image: Yes" : "Image: No");
    return parts.join(" | ");
  }

  function updateKikuCardSelection(): void {
    ctx.dom.kikuCard1.classList.toggle(
      "active",
      ctx.state.kikuSelectedCard === 1,
    );
    ctx.dom.kikuCard2.classList.toggle(
      "active",
      ctx.state.kikuSelectedCard === 2,
    );
  }

  function setKikuModalStep(step: "select" | "preview"): void {
    ctx.state.kikuModalStep = step;
    const isSelect = step === "select";
    ctx.dom.kikuSelectionStep.classList.toggle("hidden", !isSelect);
    ctx.dom.kikuPreviewStep.classList.toggle("hidden", isSelect);
    ctx.dom.kikuHint.textContent = isSelect
      ? "Press 1 or 2 to select · Enter to continue · Esc to cancel"
      : "Enter to confirm merge · Backspace to go back · Esc to cancel";
  }

  function updateKikuPreviewToggle(): void {
    ctx.dom.kikuPreviewCompactButton.classList.toggle(
      "active",
      ctx.state.kikuPreviewMode === "compact",
    );
    ctx.dom.kikuPreviewFullButton.classList.toggle(
      "active",
      ctx.state.kikuPreviewMode === "full",
    );
  }

  function renderKikuPreview(): void {
    const payload =
      ctx.state.kikuPreviewMode === "compact"
        ? ctx.state.kikuPreviewCompactData
        : ctx.state.kikuPreviewFullData;
    ctx.dom.kikuPreviewJson.textContent = payload
      ? JSON.stringify(payload, null, 2)
      : "{}";
    updateKikuPreviewToggle();
  }

  function setKikuPreviewError(message: string | null): void {
    if (!message) {
      ctx.dom.kikuPreviewError.textContent = "";
      ctx.dom.kikuPreviewError.classList.add("hidden");
      return;
    }

    ctx.dom.kikuPreviewError.textContent = message;
    ctx.dom.kikuPreviewError.classList.remove("hidden");
  }

  function openKikuFieldGroupingModal(data: {
    original: KikuDuplicateCardInfo;
    duplicate: KikuDuplicateCardInfo;
  }): void {
    if (ctx.platform.isInvisibleLayer) return;
    if (ctx.state.kikuModalOpen) return;

    ctx.state.kikuModalOpen = true;
    ctx.state.kikuOriginalData = data.original;
    ctx.state.kikuDuplicateData = data.duplicate;
    ctx.state.kikuSelectedCard = 1;

    ctx.dom.kikuCard1Expression.textContent = data.original.expression;
    ctx.dom.kikuCard1Sentence.textContent =
      data.original.sentencePreview || "(no sentence)";
    ctx.dom.kikuCard1Meta.textContent = formatMediaMeta(data.original);

    ctx.dom.kikuCard2Expression.textContent = data.duplicate.expression;
    ctx.dom.kikuCard2Sentence.textContent =
      data.duplicate.sentencePreview || "(current subtitle)";
    ctx.dom.kikuCard2Meta.textContent = formatMediaMeta(data.duplicate);

    ctx.dom.kikuDeleteDuplicateCheckbox.checked = true;
    ctx.state.kikuPendingChoice = null;
    ctx.state.kikuPreviewCompactData = null;
    ctx.state.kikuPreviewFullData = null;
    ctx.state.kikuPreviewMode = "compact";

    renderKikuPreview();
    setKikuPreviewError(null);
    setKikuModalStep("select");
    updateKikuCardSelection();

    options.syncSettingsModalSubtitleSuppression();
    ctx.dom.overlay.classList.add("interactive");
    ctx.dom.kikuModal.classList.remove("hidden");
    ctx.dom.kikuModal.setAttribute("aria-hidden", "false");
  }

  function closeKikuFieldGroupingModal(): void {
    if (!ctx.state.kikuModalOpen) return;

    ctx.state.kikuModalOpen = false;
    options.syncSettingsModalSubtitleSuppression();

    ctx.dom.kikuModal.classList.add("hidden");
    ctx.dom.kikuModal.setAttribute("aria-hidden", "true");

    setKikuPreviewError(null);
    ctx.dom.kikuPreviewJson.textContent = "";

    ctx.state.kikuPendingChoice = null;
    ctx.state.kikuPreviewCompactData = null;
    ctx.state.kikuPreviewFullData = null;
    ctx.state.kikuPreviewMode = "compact";
    setKikuModalStep("select");
    ctx.state.kikuOriginalData = null;
    ctx.state.kikuDuplicateData = null;

    if (
      !ctx.state.isOverSubtitle &&
      !options.modalStateReader.isAnyModalOpen()
    ) {
      ctx.dom.overlay.classList.remove("interactive");
    }
  }

  async function confirmKikuSelection(): Promise<void> {
    if (!ctx.state.kikuOriginalData || !ctx.state.kikuDuplicateData) return;

    const keepData =
      ctx.state.kikuSelectedCard === 1
        ? ctx.state.kikuOriginalData
        : ctx.state.kikuDuplicateData;
    const deleteData =
      ctx.state.kikuSelectedCard === 1
        ? ctx.state.kikuDuplicateData
        : ctx.state.kikuOriginalData;

    const choice: KikuFieldGroupingChoice = {
      keepNoteId: keepData.noteId,
      deleteNoteId: deleteData.noteId,
      deleteDuplicate: ctx.dom.kikuDeleteDuplicateCheckbox.checked,
      cancelled: false,
    };

    ctx.state.kikuPendingChoice = choice;
    setKikuPreviewError(null);
    ctx.dom.kikuConfirmButton.disabled = true;

    try {
      const preview: KikuMergePreviewResponse =
        await window.electronAPI.kikuBuildMergePreview({
          keepNoteId: choice.keepNoteId,
          deleteNoteId: choice.deleteNoteId,
          deleteDuplicate: choice.deleteDuplicate,
        });

      if (!preview.ok) {
        setKikuPreviewError(preview.error || "Failed to build merge preview");
        return;
      }

      ctx.state.kikuPreviewCompactData = preview.compact || {};
      ctx.state.kikuPreviewFullData = preview.full || {};
      ctx.state.kikuPreviewMode = "compact";
      renderKikuPreview();
      setKikuModalStep("preview");
    } finally {
      ctx.dom.kikuConfirmButton.disabled = false;
    }
  }

  function confirmKikuMerge(): void {
    if (!ctx.state.kikuPendingChoice) return;
    window.electronAPI.kikuFieldGroupingRespond(ctx.state.kikuPendingChoice);
    closeKikuFieldGroupingModal();
  }

  function goBackFromKikuPreview(): void {
    setKikuPreviewError(null);
    setKikuModalStep("select");
  }

  function cancelKikuFieldGrouping(): void {
    const choice: KikuFieldGroupingChoice = {
      keepNoteId: 0,
      deleteNoteId: 0,
      deleteDuplicate: true,
      cancelled: true,
    };

    window.electronAPI.kikuFieldGroupingRespond(choice);
    closeKikuFieldGroupingModal();
  }

  function handleKikuKeydown(e: KeyboardEvent): boolean {
    if (ctx.state.kikuModalStep === "preview") {
      if (e.key === "Escape") {
        e.preventDefault();
        cancelKikuFieldGrouping();
        return true;
      }
      if (e.key === "Backspace") {
        e.preventDefault();
        goBackFromKikuPreview();
        return true;
      }
      if (e.key === "Enter") {
        e.preventDefault();
        confirmKikuMerge();
        return true;
      }
      return true;
    }

    if (e.key === "Escape") {
      e.preventDefault();
      cancelKikuFieldGrouping();
      return true;
    }

    if (e.key === "1") {
      e.preventDefault();
      ctx.state.kikuSelectedCard = 1;
      updateKikuCardSelection();
      return true;
    }

    if (e.key === "2") {
      e.preventDefault();
      ctx.state.kikuSelectedCard = 2;
      updateKikuCardSelection();
      return true;
    }

    if (e.key === "ArrowLeft" || e.key === "ArrowRight") {
      e.preventDefault();
      ctx.state.kikuSelectedCard = ctx.state.kikuSelectedCard === 1 ? 2 : 1;
      updateKikuCardSelection();
      return true;
    }

    if (e.key === "Enter") {
      e.preventDefault();
      void confirmKikuSelection();
      return true;
    }

    return true;
  }

  function wireDomEvents(): void {
    ctx.dom.kikuCard1.addEventListener("click", () => {
      ctx.state.kikuSelectedCard = 1;
      updateKikuCardSelection();
    });
    ctx.dom.kikuCard1.addEventListener("dblclick", () => {
      ctx.state.kikuSelectedCard = 1;
      void confirmKikuSelection();
    });

    ctx.dom.kikuCard2.addEventListener("click", () => {
      ctx.state.kikuSelectedCard = 2;
      updateKikuCardSelection();
    });
    ctx.dom.kikuCard2.addEventListener("dblclick", () => {
      ctx.state.kikuSelectedCard = 2;
      void confirmKikuSelection();
    });

    ctx.dom.kikuConfirmButton.addEventListener("click", () => {
      void confirmKikuSelection();
    });
    ctx.dom.kikuCancelButton.addEventListener("click", () => {
      cancelKikuFieldGrouping();
    });
    ctx.dom.kikuBackButton.addEventListener("click", () => {
      goBackFromKikuPreview();
    });
    ctx.dom.kikuFinalConfirmButton.addEventListener("click", () => {
      confirmKikuMerge();
    });
    ctx.dom.kikuFinalCancelButton.addEventListener("click", () => {
      cancelKikuFieldGrouping();
    });

    ctx.dom.kikuPreviewCompactButton.addEventListener("click", () => {
      ctx.state.kikuPreviewMode = "compact";
      renderKikuPreview();
    });
    ctx.dom.kikuPreviewFullButton.addEventListener("click", () => {
      ctx.state.kikuPreviewMode = "full";
      renderKikuPreview();
    });
  }

  return {
    closeKikuFieldGroupingModal,
    handleKikuKeydown,
    openKikuFieldGroupingModal,
    wireDomEvents,
  };
}
