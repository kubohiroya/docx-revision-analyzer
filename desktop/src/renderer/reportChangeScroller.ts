import { el } from "./extensions";

const ROW_HEIGHT = 64;
const OVERSCAN = 5;
const VIRTUALIZE_AFTER = 24;

export function reportChangeScroller<T>(
  scroller: HTMLElement, list: HTMLUListElement, render: (item: T) => HTMLLIElement,
) {
  let items: readonly T[] = [];
  let first = -1;
  let last = -1;
  list.style.setProperty("--report-change-row-height", `${ROW_HEIGHT}px`);

  function spacer(height: number): HTMLLIElement {
    const item = el("li", undefined, "report-change-spacer");
    item.setAttribute("aria-hidden", "true");
    item.style.height = `${height}px`;
    return item;
  }

  function draw(): void {
    const virtual = items.length > VIRTUALIZE_AFTER;
    const start = virtual ? Math.max(0, Math.floor(scroller.scrollTop / ROW_HEIGHT) - OVERSCAN) : 0;
    const end = virtual ? Math.min(items.length, Math.ceil((scroller.scrollTop + scroller.clientHeight) / ROW_HEIGHT) + OVERSCAN)
      : items.length;
    if (start === first && end === last) return;
    first = start;
    last = end;
    list.classList.toggle("report-virtual-changes", virtual);
    const nodes = [];
    if (start) nodes.push(spacer(start * ROW_HEIGHT));
    for (let i = start; i < end; i++) {
      const row = render(items[i]);
      row.dataset.changeIndex = String(i);
      row.setAttribute("aria-posinset", String(i + 1));
      row.setAttribute("aria-setsize", String(items.length));
      nodes.push(row);
    }
    if (end < items.length) nodes.push(spacer((items.length - end) * ROW_HEIGHT));
    const active = document.activeElement;
    const focusedRow = active?.closest<HTMLElement>("[data-change-index]");
    const restore = focusedRow && list.contains(focusedRow);
    list.replaceChildren(...nodes);
    if (restore) {
      const control = active instanceof HTMLInputElement ? "input" : "button";
      const item = list.querySelector<HTMLElement>(`[data-change-index="${focusedRow.dataset.changeIndex}"] ${control}`);
      (item ?? scroller).focus({ preventScroll: true });
    }
  }

  scroller.onscroll = draw;
  const measure = new ResizeObserver(() => {
    if (scroller.isConnected) draw();
    else measure.disconnect();
  });
  measure.observe(scroller);
  scroller.onkeydown = (event) => {
    if (!(event.target instanceof Element) || !items.length) return;
    const current = event.target.closest<HTMLElement>("[data-change-index]");
    if (!current && event.target !== scroller) return;
    const index = current ? Number(current.dataset.changeIndex) : -1;
    let next: number;
    switch (event.key) {
      case "ArrowDown": next = Math.min(items.length - 1, index + 1); break;
      case "ArrowUp": next = Math.max(0, index - 1); break;
      case "Home": next = 0; break;
      case "End": next = items.length - 1; break;
      default: return;
    }
    event.preventDefault();
    if (items.length > VIRTUALIZE_AFTER) {
      if (next * ROW_HEIGHT < scroller.scrollTop) scroller.scrollTop = next * ROW_HEIGHT;
      else if ((next + 1) * ROW_HEIGHT > scroller.scrollTop + scroller.clientHeight) {
        scroller.scrollTop = (next + 1) * ROW_HEIGHT - scroller.clientHeight;
      }
      draw();
    }
    list.querySelector<HTMLButtonElement>(`[data-change-index="${next}"] button`)?.focus();
  };

  return {
    setItems(next: readonly T[], reset: boolean) {
      items = next;
      if (reset) scroller.scrollTop = 0;
      first = last = -1;
      draw();
    },
  };
}
