import { CompactFileList } from "../common/CompactFileList";
import { LayerControlPanel } from "../layer-control/LayerControlPanel";
import { LayerPreviewPanel } from "../layer-control/LayerPreviewPanel";
import { usePsdStore } from "../../store/psdStore";
import { DropZone } from "../file-browser/DropZone";
import { useOpenInPhotoshop } from "../../hooks/useOpenInPhotoshop";
import { TextExtractButton } from "../common/TextExtractButton";

export function LayerControlView() {
  const files = usePsdStore((state) => state.files);
  const hasFiles = files.length > 0;
  const { openFileInPhotoshop } = useOpenInPhotoshop();

  if (!hasFiles) {
    return <DropZone />;
  }

  return (
    <div className="flex h-full overflow-hidden" data-tool-panel>
      {/* File List */}
      <CompactFileList className="w-52 flex-shrink-0 border-r border-border" />

      {/* Settings */}
      <div className="w-[360px] flex-shrink-0 border-r border-border overflow-hidden">
        <LayerControlPanel />
      </div>

      {/* Layer Preview */}
      <div className="flex-1 overflow-hidden relative">
        <LayerPreviewPanel onOpenInPhotoshop={openFileInPhotoshop} />
        {/* テキスト抽出ボタン（常時表示） */}
        <div className="absolute bottom-6 right-6 flex flex-col items-end gap-3 z-10">
          <TextExtractButton />
        </div>
      </div>
    </div>
  );
}
