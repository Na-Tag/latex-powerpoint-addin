const TAG_LATEX = "LATEX_EQUATION_SOURCE";
const TAG_META = "LATEX_EQUATION_META";
const APP_TITLE = "LaTeX Equation";
const DEFAULT_INSERT = { left: 72, top: 120 };

const els = {};
let lastRendered = null;
let selectionTimer = null;
let isOfficeReady = false;
let isMutatingShape = false;

function $(id) {
  return document.getElementById(id);
}

function setStatus(message, kind = "") {
  const status = els.status;
  status.textContent = message;
  status.className = `status ${kind}`.trim();
}

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

function getFormState() {
  const latex = els.latexInput.value.trim();
  const heightPt = Math.max(12, Math.min(180, Number(els.heightPt.value || 42)));
  const color = els.equationColor.value || "#111111";
  const display = els.displayMode.checked;
  return { latex, heightPt, color, display };
}

function waitForMathJax() {
  return new Promise((resolve, reject) => {
    const started = Date.now();
    const tick = () => {
      if (window.MathJax && MathJax.startup && MathJax.startup.promise) {
        MathJax.startup.promise.then(resolve).catch(reject);
        return;
      }
      if (Date.now() - started > 10000) {
        reject(new Error("MathJaxの読み込みがタイムアウトしました。"));
        return;
      }
      setTimeout(tick, 50);
    };
    tick();
  });
}

async function renderLatexToSvg({ latex, color, display }) {
  if (!latex) throw new Error("LaTeXが空です。");
  await waitForMathJax();
  const wrapper = await MathJax.tex2svgPromise(latex, { display });
  const svg = wrapper.querySelector("svg");
  if (!svg) throw new Error("SVGへの変換に失敗しました。LaTeXを確認してください。");

  svg.setAttribute("xmlns", "http://www.w3.org/2000/svg");
  svg.setAttribute("role", "img");
  svg.setAttribute("aria-label", latex);
  svg.setAttribute("style", `color:${color};`);

  // Office.CoercionType.XmlSvg expects XML data containing an SVG image.
  const svgText = new XMLSerializer().serializeToString(svg);
  const viewBox = svg.getAttribute("viewBox") || "0 0 1000 300";
  const [, , rawWidth, rawHeight] = viewBox.split(/\s+/).map(Number);
  const aspectRatio = rawWidth && rawHeight ? rawWidth / rawHeight : 3;
  return { svgText, aspectRatio };
}

async function updatePreview() {
  try {
    const state = getFormState();
    const rendered = await renderLatexToSvg(state);
    lastRendered = { ...rendered, ...state };
    els.preview.innerHTML = rendered.svgText;
    setStatus("プレビューを更新しました。", "ok");
    return lastRendered;
  } catch (error) {
    els.preview.innerHTML = `<span class="status error">${escapeHtml(error.message)}</span>`;
    setStatus(error.message, "error");
    throw error;
  }
}

async function getRenderedEquation() {
  const state = getFormState();
  if (!lastRendered || lastRendered.latex !== state.latex || lastRendered.color !== state.color || lastRendered.display !== state.display || lastRendered.heightPt !== state.heightPt) {
    return updatePreview();
  }
  return lastRendered;
}

function setSelectedDataAsync(data, options) {
  return new Promise((resolve, reject) => {
    Office.context.document.setSelectedDataAsync(data, options, (result) => {
      if (result.status === Office.AsyncResultStatus.Failed) {
        reject(new Error(result.error.message));
      } else {
        resolve(result.value);
      }
    });
  });
}

function delay(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function asFiniteNumber(value, fallback = 0) {
  const number = Number(value);
  return Number.isFinite(number) ? number : fallback;
}

function shapeDistance(shape, target) {
  return Math.abs(asFiniteNumber(shape.left) - asFiniteNumber(target.left))
    + Math.abs(asFiniteNumber(shape.top) - asFiniteNumber(target.top))
    + Math.abs(asFiniteNumber(shape.width) - asFiniteNumber(target.width)) * 0.25
    + Math.abs(asFiniteNumber(shape.height) - asFiniteNumber(target.height)) * 0.25;
}

async function getSelectedSlideShapeSnapshots() {
  return PowerPoint.run(async (context) => {
    const slide = context.presentation.getSelectedSlides().getItemAt(0);
    const shapes = slide.shapes;
    shapes.load("items");
    await context.sync();

    shapes.items.forEach((shape) => {
      shape.load("id,name,left,top,width,height,zOrderPosition,altTextTitle,altTextDescription,type");
    });
    await context.sync();

    return shapes.items.map((shape, index) => ({
      id: shape.id,
      index,
      name: shape.name,
      left: shape.left,
      top: shape.top,
      width: shape.width,
      height: shape.height,
      zOrderPosition: shape.zOrderPosition,
      altTextTitle: shape.altTextTitle,
      altTextDescription: shape.altTextDescription,
      type: shape.type
    }));
  });
}

function findBestShapeAfterInsert(beforeShapes, afterShapes, target, previousSelectedId) {
  const beforeIds = new Set(beforeShapes.map((shape) => shape.id));
  const newShapes = afterShapes.filter((shape) => !beforeIds.has(shape.id));

  const chooseClosest = (candidates) => {
    if (!candidates.length) return null;
    return candidates
      .map((shape) => ({
        shape,
        score: shapeDistance(shape, target) - asFiniteNumber(shape.zOrderPosition) * 0.01
      }))
      .sort((a, b) => a.score - b.score)[0].shape;
  };

  // Most reliable path: Common API inserted a new SVG image shape.
  const inserted = chooseClosest(newShapes);
  if (inserted) return inserted;

  // Some PowerPoint builds replace the selected image in-place and preserve the shape id.
  if (previousSelectedId) {
    const sameId = afterShapes.find((shape) => shape.id === previousSelectedId);
    if (sameId) return sameId;
  }

  // Last-resort fallback: find a shape that now sits where the equation should be.
  const nearTarget = afterShapes.filter((shape) => shapeDistance(shape, target) < 12);
  return chooseClosest(nearTarget) || chooseClosest(afterShapes);
}

async function tagShapeById(shapeId, metadata) {
  return PowerPoint.run(async (context) => {
    const slide = context.presentation.getSelectedSlides().getItemAt(0);
    const shape = slide.shapes.getItem(shapeId);
    shape.name = APP_TITLE;
    shape.altTextTitle = APP_TITLE;
    shape.altTextDescription = JSON.stringify(metadata);
    shape.tags.add(TAG_LATEX, metadata.latex);
    shape.tags.add(TAG_META, JSON.stringify(metadata));
    await context.sync();
  });
}

async function tagInsertedOrReplacedShape({ beforeShapes, target, previousSelectedId, metadata }) {
  // Give PowerPoint a tiny moment to commit the Common API insertion before reading the shape tree.
  await delay(120);
  const afterShapes = await getSelectedSlideShapeSnapshots();
  const shape = findBestShapeAfterInsert(beforeShapes, afterShapes, target, previousSelectedId);
  if (!shape || !shape.id) {
    throw new Error("挿入された数式図形を特定できませんでした。もう一度選択して更新してください。");
  }
  await tagShapeById(shape.id, metadata);
  return shape.id;
}

async function deleteShapeById(shapeId) {
  if (!shapeId) return;
  await PowerPoint.run(async (context) => {
    const slide = context.presentation.getSelectedSlides().getItemAt(0);
    const shape = slide.shapes.getItem(shapeId);
    shape.delete();
    await context.sync();
  });
}

async function getSingleSelectedShapeId() {
  if (!isOfficeReady) return null;
  return PowerPoint.run(async (context) => {
    const shapes = context.presentation.getSelectedShapes();
    const count = shapes.getCount();
    await context.sync();
    if (count.value !== 1) return null;
    const shape = shapes.getItemAt(0);
    shape.load("id");
    await context.sync();
    return shape.id;
  });
}

async function getSelectedEquationInfo() {
  if (!isOfficeReady) return null;
  return PowerPoint.run(async (context) => {
    const shapes = context.presentation.getSelectedShapes();
    const count = shapes.getCount();
    await context.sync();
    if (count.value !== 1) return null;

    const shape = shapes.getItemAt(0);
    shape.load("id,name,left,top,width,height,altTextTitle,altTextDescription");
    const latexTag = shape.tags.getItemOrNullObject(TAG_LATEX);
    latexTag.load("key,value");
    const metaTag = shape.tags.getItemOrNullObject(TAG_META);
    metaTag.load("key,value");
    await context.sync();

    let latex = null;
    let meta = null;
    if (!latexTag.isNullObject) latex = latexTag.value;
    if (!metaTag.isNullObject) {
      try { meta = JSON.parse(metaTag.value); } catch { meta = null; }
    }
    if (!latex && shape.altTextDescription) {
      try {
        const parsed = JSON.parse(shape.altTextDescription);
        if (parsed && parsed.latex) {
          latex = parsed.latex;
          meta = parsed;
        }
      } catch { /* ignore legacy/non-json alt text */ }
    }

    const isEquation = Boolean(latex) || shape.name === APP_TITLE || shape.altTextTitle === APP_TITLE;
    return {
      id: shape.id,
      latex,
      meta,
      isEquation,
      left: shape.left,
      top: shape.top,
      width: shape.width,
      height: shape.height
    };
  });
}

async function tagSelectedOrLastInserted(metadata) {
  const beforeShapes = await getSelectedSlideShapeSnapshots();
  await tagInsertedOrReplacedShape({
    beforeShapes,
    target: DEFAULT_INSERT,
    previousSelectedId: null,
    metadata
  });
}

async function insertEquation({ replaceSelected }) {
  if (isMutatingShape) return;
  isMutatingShape = true;
  try {
    const rendered = await getRenderedEquation();
    let selected = null;
    let left = DEFAULT_INSERT.left;
    let top = DEFAULT_INSERT.top;
    let height = rendered.heightPt;
    let width = Math.max(20, rendered.heightPt * rendered.aspectRatio);
    let previousSelectedId = null;

    if (replaceSelected) {
      selected = await getSelectedEquationInfo();
      if (!selected || !selected.isEquation) {
        throw new Error("更新するには、このアドインで作成した数式を1つ選択してください。");
      }
      previousSelectedId = selected.id;
      left = selected.left;
      top = selected.top;
      height = selected.height || rendered.heightPt;
      width = Math.max(20, height * rendered.aspectRatio);
    } else {
      // If PowerPoint treats the current selection as replacement target,
      // this lets us still tag the overwritten image correctly.
      previousSelectedId = await getSingleSelectedShapeId();
    }

    const metadata = {
      app: "latex-powerpoint-equation-addin",
      version: 2,
      latex: rendered.latex,
      display: rendered.display,
      color: rendered.color,
      heightPt: rendered.heightPt,
      updatedAt: new Date().toISOString()
    };

    const beforeShapes = await getSelectedSlideShapeSnapshots();
    const target = { left, top, width, height };

    await setSelectedDataAsync(rendered.svgText, {
      coercionType: Office.CoercionType.XmlSvg,
      imageLeft: left,
      imageTop: top,
      imageWidth: width,
      imageHeight: height
    });

    // Do not trust the post-insertion selection. PowerPoint may keep the old object selected,
    // select the new object, or briefly select nothing depending on platform/build.
    const newShapeId = await tagInsertedOrReplacedShape({ beforeShapes, target, previousSelectedId, metadata });

    // In update mode, keep only the newly rendered equation.
    // Some PowerPoint builds insert a new SVG instead of replacing the selected one,
    // so explicitly remove the previous equation shape when the new shape has a different id.
    if (replaceSelected && previousSelectedId && newShapeId !== previousSelectedId) {
      await deleteShapeById(previousSelectedId);
    }

    setStatus(replaceSelected ? "選択した数式を更新しました。古い数式は削除しました。" : "数式を挿入しました。", "ok");
  } catch (error) {
    setStatus(error.message, "error");
  } finally {
    isMutatingShape = false;
  }
}

async function loadSelectedEquation({ quiet = false } = {}) {
  try {
    const selected = await getSelectedEquationInfo();
    if (!selected || !selected.isEquation || !selected.latex) {
      if (!quiet) setStatus("このアドインで作成した数式を1つ選択してください。", "error");
      return;
    }
    els.latexInput.value = selected.latex;
    if (selected.meta) {
      if (typeof selected.meta.display === "boolean") els.displayMode.checked = selected.meta.display;
      if (selected.meta.color) els.equationColor.value = selected.meta.color;
      if (selected.meta.heightPt) els.heightPt.value = selected.meta.heightPt;
    }
    await updatePreview();
    setStatus("選択した数式のLaTeXを読み込みました。", "ok");
  } catch (error) {
    if (!quiet) setStatus(error.message, "error");
  }
}

function wireEvents() {
  els.previewButton.addEventListener("click", () => updatePreview().catch(() => {}));
  els.insertButton.addEventListener("click", () => insertEquation({ replaceSelected: false }));
  els.loadButton.addEventListener("click", () => loadSelectedEquation());
  els.updateButton.addEventListener("click", () => insertEquation({ replaceSelected: true }));

  els.latexInput.addEventListener("input", () => {
    clearTimeout(selectionTimer);
    selectionTimer = setTimeout(() => updatePreview().catch(() => {}), 350);
  });
  els.displayMode.addEventListener("change", () => updatePreview().catch(() => {}));
  els.equationColor.addEventListener("input", () => updatePreview().catch(() => {}));

  Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, () => {
    if (isMutatingShape) return;
    clearTimeout(selectionTimer);
    selectionTimer = setTimeout(() => {
      if (!isMutatingShape) loadSelectedEquation({ quiet: true });
    }, 250);
  });
}

Office.onReady((info) => {
  els.latexInput = $("latexInput");
  els.displayMode = $("displayMode");
  els.heightPt = $("heightPt");
  els.equationColor = $("equationColor");
  els.preview = $("preview");
  els.status = $("status");
  els.previewButton = $("previewButton");
  els.insertButton = $("insertButton");
  els.loadButton = $("loadButton");
  els.updateButton = $("updateButton");

  if (info.host !== Office.HostType.PowerPoint) {
    setStatus("PowerPointでこのアドインを開いてください。", "error");
    return;
  }

  isOfficeReady = true;
  wireEvents();
  updatePreview()
    .then(() => setStatus("準備完了です。", "ok"))
    .catch(() => {});
});
