import { el } from "./extensions";

type Panel = { body: HTMLElement; label: HTMLElement };
const MIN_HEIGHT = 72;

export function reportAccordionSplitters(
  panels: readonly Panel[], heights: (number | undefined)[],
  label: (above: string, below: string) => string,
) {
  const handles = panels.slice(0, -1).map((above, index) => {
    const below = panels[index + 1];
    const handle = el("div", undefined, "report-accordion-splitter");
    const name = label(above.label.textContent ?? "", below.label.textContent ?? "");
    handle.setAttribute("role", "separator");
    handle.setAttribute("aria-orientation", "horizontal");
    handle.setAttribute("aria-label", name);
    handle.setAttribute("aria-controls", `${above.body.id} ${below.body.id}`);
    handle.title = name;
    handle.tabIndex = 0;

    function dimensions() {
      const top = above.body.getBoundingClientRect().height;
      const total = top + below.body.getBoundingClientRect().height;
      const minimum = Math.min(MIN_HEIGHT, Math.floor(total / 2));
      return { top, total, minimum };
    }
    function resize(top: number, total: number, minimum: number) {
      heights[index] = Math.max(minimum, Math.min(total - minimum, top));
      heights[index + 1] = total - heights[index];
      update();
    }
    handle.onpointerdown = (event) => {
      if (!event.isPrimary || event.button !== 0) return;
      handle.focus({ preventScroll: true });
      event.preventDefault();
      handle.setPointerCapture(event.pointerId);
      handle.classList.add("resizing");
      const { top, total, minimum } = dimensions();
      const { clientY: start, pointerId } = event;
      const controller = new AbortController();
      const options = { capture: true, signal: controller.signal };
      const finish = () => {
        controller.abort();
        handle.classList.remove("resizing");
      };
      window.addEventListener("pointermove", (move) => {
        if (move.pointerId === pointerId) resize(top + move.clientY - start, total, minimum);
      }, options);
      const released = (release: PointerEvent) => {
        if (release.pointerId === pointerId) finish();
      };
      window.addEventListener("pointerup", released, options);
      window.addEventListener("pointercancel", released, options);
      window.addEventListener("blur", finish, options);
    };
    handle.onkeydown = (event) => {
      const { top, total, minimum } = dimensions();
      let next: number;
      switch (event.key) {
        case "ArrowUp": next = top - 16; break;
        case "ArrowDown": next = top + 16; break;
        case "Home": next = minimum; break;
        case "End": next = total - minimum; break;
        default: return;
      }
      event.preventDefault();
      resize(next, total, minimum);
    };
    handle.ondblclick = () => {
      heights[index] = heights[index + 1] = undefined;
      update();
    };
    return { handle, above, below, dimensions };
  });

  function update() {
    for (const [index, { body }] of panels.entries()) {
      const height = heights[index];
      body.classList.toggle("report-accordion-sized", height !== undefined);
      if (height === undefined) body.style.removeProperty("--report-accordion-height");
      else body.style.setProperty("--report-accordion-height", `${height}px`);
    }
    for (const { handle, above, below, dimensions } of handles) {
      handle.hidden = above.body.hidden || below.body.hidden;
      if (handle.hidden) continue;
      const { top, total, minimum } = dimensions();
      const percent = (height: number) => String(total > 0 ? Math.round(100 * height / total) : 50);
      handle.setAttribute("aria-valuemin", percent(minimum));
      handle.setAttribute("aria-valuemax", percent(total - minimum));
      handle.setAttribute("aria-valuenow", percent(top));
    }
  }

  return { handles: handles.map(({ handle }) => handle), update };
}
