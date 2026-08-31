import {
  useState,
  useCallback,
  useRef,
  type MouseEventHandler,
  type MouseEvent,
  type TouchEventHandler,
} from "react";

const useLongPress = (
  onLongPress: MouseEventHandler | TouchEventHandler,
  onClick?: MouseEventHandler | TouchEventHandler,
  { shouldPreventDefault = true, delay = 300 } = {},
) => {
  const [longPressTriggered, setLongPressTriggered] = useState(false);
  const timeout = useRef<NodeJS.Timeout>(null);
  const target = useRef<EventTarget>(null);

  const start = useCallback(
    (event: MouseEvent | TouchEvent) => {
      if (shouldPreventDefault && event.target) {
        event.target.addEventListener("touchend", preventDefault, {
          passive: false,
        });
        target.current = event.target;
      }

      timeout.current = setTimeout(() => {
        onLongPress(event);
        setLongPressTriggered(true);
      }, delay);
    },
    [onLongPress, delay, shouldPreventDefault],
  );

  const clear = useCallback(
    (event: MouseEvent | TouchEvent, shouldTriggerClick = true) => {
      timeout.current && clearTimeout(timeout.current);
      onClick && shouldTriggerClick && !longPressTriggered && onClick(event);
      setLongPressTriggered(false);

      if (shouldPreventDefault && target.current) {
        target.current.removeEventListener("touchend", preventDefault);
      }
    },
    [shouldPreventDefault, onClick, longPressTriggered],
  );

  return {
    onMouseDown: (e: MouseEvent) => start(e),
    onTouchStart: (e: TouchEvent) => start(e),
    onMouseUp: (e: MouseEvent) => clear(e),
    onMouseLeave: (e: MouseEvent) => clear(e, false),
    onTouchEnd: (e: TouchEvent) => clear(e, false),
    onTouchMove: (e: TouchEvent) => clear(e, false),
  };
};

// Helper to prevent default touch actions
const preventDefault = (event: Event) => {
  if (!("touches" in event)) return;
  if (!Array.isArray(event.touches)) return;
  if (event.touches.length < 2 && event.preventDefault) {
    if (event.cancelable) event.preventDefault();
  }
};

export default useLongPress;
