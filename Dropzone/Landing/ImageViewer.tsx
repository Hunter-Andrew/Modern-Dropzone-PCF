import * as React from "react";
import { IconButton, IButtonStyles, Modal, Stack } from "@fluentui/react";

// Lightweight image preview built on the platform-provided Fluent UI, so the
// control does not have to bundle its own copy of react-dom.

export interface ImageViewerProps {
  src: string;
  alt: string;
  onClose: () => void;
}

const MIN_SCALE = 0.1;
const MAX_SCALE = 10;
const ZOOM_STEP = 1.25;

const toolbarButtonStyles: IButtonStyles = {
  root: { color: "#fff", width: 40, height: 40 },
  rootHovered: { color: "#fff", backgroundColor: "rgba(255,255,255,0.15)" },
  rootPressed: { color: "#fff", backgroundColor: "rgba(255,255,255,0.25)" },
  rootDisabled: { color: "rgba(255,255,255,0.35)", backgroundColor: "transparent" },
  icon: { fontSize: 18 },
};

const clamp = (value: number) => Math.min(MAX_SCALE, Math.max(MIN_SCALE, value));

export const ImageViewer: React.FC<ImageViewerProps> = ({ src, alt, onClose }) => {
  const [scale, setScale] = React.useState(1);
  const [rotation, setRotation] = React.useState(0);
  const [offset, setOffset] = React.useState({ x: 0, y: 0 });
  const [dragging, setDragging] = React.useState(false);
  const drag = React.useRef<{
    startX: number;
    startY: number;
    originX: number;
    originY: number;
  } | null>(null);

  const reset = () => {
    setScale(1);
    setRotation(0);
    setOffset({ x: 0, y: 0 });
  };

  const onWheel = (e: React.WheelEvent) => {
    setScale((s) => clamp(e.deltaY < 0 ? s * ZOOM_STEP : s / ZOOM_STEP));
  };

  const onPointerDown = (e: React.PointerEvent) => {
    if (e.button !== 0) return;
    e.preventDefault();
    (e.currentTarget as HTMLElement).setPointerCapture(e.pointerId);
    drag.current = { startX: e.clientX, startY: e.clientY, originX: offset.x, originY: offset.y };
    setDragging(true);
  };

  const onPointerMove = (e: React.PointerEvent) => {
    const d = drag.current;
    if (!d) return;
    setOffset({ x: d.originX + e.clientX - d.startX, y: d.originY + e.clientY - d.startY });
  };

  const endDrag = () => {
    drag.current = null;
    setDragging(false);
  };

  return (
    <Modal
      isOpen
      onDismiss={onClose}
      isBlocking={false}
      styles={{
        main: {
          width: "100%",
          height: "100%",
          maxWidth: "100%",
          maxHeight: "100%",
          backgroundColor: "transparent",
          boxShadow: "none",
        },
        scrollableContent: { height: "100%", overflow: "hidden" },
      }}
      overlay={{ styles: { root: { backgroundColor: "rgba(0,0,0,0.85)" } } }}
    >
      <div
        style={{
          position: "relative",
          width: "100%",
          height: "100%",
          overflow: "hidden",
          touchAction: "none",
        }}
        onWheel={onWheel}
        onPointerDown={onPointerDown}
        onPointerMove={onPointerMove}
        onPointerUp={endDrag}
        onPointerCancel={endDrag}
        onDoubleClick={reset}
      >
        <img
          src={src}
          alt={alt}
          draggable={false}
          style={{
            position: "absolute",
            top: "50%",
            left: "50%",
            maxWidth: "90%",
            maxHeight: "85%",
            transform: `translate(-50%, -50%) translate(${offset.x}px, ${offset.y}px) rotate(${rotation}deg) scale(${scale})`,
            transition: dragging ? "none" : "transform 0.15s ease-out",
            cursor: dragging ? "grabbing" : "grab",
            userSelect: "none",
          }}
        />
      </div>

      <IconButton
        iconProps={{ iconName: "Cancel" }}
        title="Close"
        ariaLabel="Close"
        onClick={onClose}
        styles={toolbarButtonStyles}
        style={{ position: "absolute", top: 12, right: 12 }}
      />

      <div
        style={{
          position: "absolute",
          bottom: 16,
          left: "50%",
          transform: "translateX(-50%)",
          textAlign: "center",
          color: "#fff",
        }}
      >
        <div style={{ marginBottom: 8, fontSize: 14, whiteSpace: "nowrap" }}>{alt}</div>
        <Stack
          horizontal
          tokens={{ childrenGap: 4 }}
          styles={{ root: { backgroundColor: "rgba(0,0,0,0.5)", borderRadius: 4, padding: 4 } }}
        >
          <IconButton
            iconProps={{ iconName: "ZoomIn" }}
            title="Zoom in"
            ariaLabel="Zoom in"
            disabled={scale >= MAX_SCALE}
            onClick={() => setScale((s) => clamp(s * ZOOM_STEP))}
            styles={toolbarButtonStyles}
          />
          <IconButton
            iconProps={{ iconName: "ZoomOut" }}
            title="Zoom out"
            ariaLabel="Zoom out"
            disabled={scale <= MIN_SCALE}
            onClick={() => setScale((s) => clamp(s / ZOOM_STEP))}
            styles={toolbarButtonStyles}
          />
          <IconButton
            iconProps={{ iconName: "Rotate90CounterClockwise" }}
            title="Rotate left"
            ariaLabel="Rotate left"
            onClick={() => setRotation((r) => r - 90)}
            styles={toolbarButtonStyles}
          />
          <IconButton
            iconProps={{ iconName: "Rotate90Clockwise" }}
            title="Rotate right"
            ariaLabel="Rotate right"
            onClick={() => setRotation((r) => r + 90)}
            styles={toolbarButtonStyles}
          />
          <IconButton
            iconProps={{ iconName: "Refresh" }}
            title="Reset"
            ariaLabel="Reset"
            onClick={reset}
            styles={toolbarButtonStyles}
          />
        </Stack>
      </div>
    </Modal>
  );
};
