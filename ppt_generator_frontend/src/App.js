import React, { useEffect, useMemo, useRef, useState } from "react";
import PptxGenJS from "pptxgenjs";
import JSZip from "jszip";
import { XMLParser } from "fast-xml-parser";
import { saveAs } from "file-saver";
import "./App.css";

/**
 * Slide preview model
 * @typedef {Object} SlidePreview
 * @property {number} index
 * @property {string} title
 * @property {string[]} bullets
 */

/**
 * @typedef {Object} SlideInput
 * @property {string} title
 * @property {string} bulletsRaw
 */

// PUBLIC_INTERFACE
function App() {
  /** Ocean Professional theme constants */
  const theme = useMemo(
    () => ({
      primary: "#2563EB",
      secondary: "#F59E0B",
      bg: "#f9fafb",
      surface: "#ffffff",
      text: "#111827",
      error: "#EF4444",
    }),
    []
  );

  const [deckTitle, setDeckTitle] = useState("Ocean Professional Pitch");
  const [author, setAuthor] = useState("Your Name");
  const [accentMode, setAccentMode] = useState("gradient"); // gradient | solid
  const [slideCount, setSlideCount] = useState(5);
  const [slides, setSlides] = useState(
    /** @type {SlideInput[]} */ (
      Array.from({ length: 5 }).map((_, i) => ({
        title:
          i === 0
            ? "Executive Summary"
            : i === 1
              ? "Problem"
              : i === 2
                ? "Solution"
                : i === 3
                  ? "Market"
                  : "Next Steps",
        bulletsRaw:
          i === 0
            ? "What we do\nWhy now\nWhat success looks like"
            : i === 1
              ? "Pain points today\nCost of inaction\nWho is affected"
              : i === 2
                ? "Key approach\nDifferentiators\nHow it works"
                : i === 3
                  ? "TAM / SAM / SOM\nCompetitive landscape\nGo-to-market"
                  : "Milestones\nTeam & needs\nCall to action",
      }))
    )
  );

  const [activeSlideIndex, setActiveSlideIndex] = useState(0);

  const [pptxBlob, setPptxBlob] = useState(null);
  const [pptxFilename, setPptxFilename] = useState("presentation.pptx");
  const [previewSlides, setPreviewSlides] = useState(/** @type {SlidePreview[]} */ ([]));

  const [status, setStatus] = useState({
    kind: "idle", // idle | generating | ready | error
    message: "Configure your deck, then generate a PPTX.",
  });

  const leftPanelRef = useRef(null);

  useEffect(() => {
    // Keep slide inputs array in sync with slideCount
    setSlides((prev) => {
      const next = [...prev];
      if (slideCount > next.length) {
        for (let i = next.length; i < slideCount; i++) {
          next.push({
            title: `Slide ${i + 1}`,
            bulletsRaw: "Point one\nPoint two\nPoint three",
          });
        }
      } else if (slideCount < next.length) {
        next.length = slideCount;
      }
      return next;
    });
  }, [slideCount]);

  useEffect(() => {
    // Ensure active index stays in range
    setActiveSlideIndex((idx) => Math.min(Math.max(idx, 0), Math.max(slideCount - 1, 0)));
  }, [slideCount]);

  const activeSlide = slides[activeSlideIndex];

  function parseBullets(bulletsRaw) {
    return bulletsRaw
      .split("\n")
      .map((s) => s.trim())
      .filter(Boolean)
      .slice(0, 8);
  }

  function buildFilename() {
    const safe = (deckTitle || "presentation")
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, "-")
      .replace(/(^-|-$)/g, "");
    return `${safe || "presentation"}.pptx`;
  }

  /**
   * Create PPTX using PptxGenJS and return as Blob.
   * - Uses a simple, professional theme with blue & amber accents.
   */
  async function generatePptxBlob() {
    const pptx = new PptxGenJS();

    // 16:9
    pptx.layout = "LAYOUT_WIDE";

    // Basic metadata
    pptx.author = author || "PPT Generator";
    pptx.company = "In-browser PPT Generator";
    pptx.subject = deckTitle || "Presentation";
    pptx.title = deckTitle || "Presentation";

    // Theme fonts: stick to common cross-platform fonts
    pptx.theme = {
      headFontFace: "Aptos Display",
      bodyFontFace: "Aptos",
      lang: "en-US",
    };

    const slideW = 13.333; // inches for wide (PptxGenJS wide)
    const slideH = 7.5;

    const bg = "FFFFFF";
    const text = theme.text.replace("#", "");
    const primary = theme.primary.replace("#", "");
    const secondary = theme.secondary.replace("#", "");

    /**
     * Draws an Ocean-style header accent.
     * Note: PPTX doesn't support CSS gradients; we emulate with layered shapes.
     */
    function addHeaderAccent(s) {
      // Base top band
      s.addShape(pptx.ShapeType.rect, {
        x: 0,
        y: 0,
        w: slideW,
        h: 0.65,
        fill: { color: primary },
        line: { color: primary },
      });

      if (accentMode === "gradient") {
        // "Gradient" accent using two transparent-ish overlays
        s.addShape(pptx.ShapeType.rect, {
          x: slideW * 0.62,
          y: 0,
          w: slideW * 0.38,
          h: 0.65,
          fill: { color: secondary, transparency: 35 },
          line: { color: secondary, transparency: 100 },
        });
        s.addShape(pptx.ShapeType.roundRect, {
          x: slideW * 0.75,
          y: 0.12,
          w: slideW * 0.22,
          h: 0.42,
          fill: { color: "FFFFFF", transparency: 78 },
          line: { color: "FFFFFF", transparency: 100 },
          radius: 0.18,
        });
      }
    }

    function addFooter(s, slideNo, total) {
      const footerY = slideH - 0.4;
      s.addText(`${slideNo}/${total}`, {
        x: slideW - 1.05,
        y: footerY,
        w: 0.9,
        h: 0.25,
        fontFace: "Aptos",
        fontSize: 10,
        color: "6B7280",
        align: "right",
      });

      s.addText(author || "", {
        x: 0.6,
        y: footerY,
        w: slideW - 2.0,
        h: 0.25,
        fontFace: "Aptos",
        fontSize: 10,
        color: "6B7280",
        align: "left",
      });
    }

    const total = slides.length;

    slides.forEach((si, idx) => {
      const s = pptx.addSlide();

      // Background
      s.background = { color: bg };

      addHeaderAccent(s);

      // Title
      const title = (si.title || `Slide ${idx + 1}`).slice(0, 80);
      s.addText(title, {
        x: 0.75,
        y: 0.85,
        w: slideW - 1.5,
        h: 0.8,
        fontFace: "Aptos Display",
        fontSize: 34,
        bold: true,
        color: text,
      });

      // Subtle divider line
      s.addShape(pptx.ShapeType.line, {
        x: 0.75,
        y: 1.8,
        w: slideW - 1.5,
        h: 0,
        line: { color: primary, transparency: 70, width: 2 },
      });

      // Bullets
      const bullets = parseBullets(si.bulletsRaw);
      const bulletText = bullets.map((b) => `• ${b}`).join("\n");
      s.addText(bulletText || "• Add bullet points in the sidebar", {
        x: 0.95,
        y: 2.2,
        w: slideW - 1.9,
        h: 4.7,
        fontFace: "Aptos",
        fontSize: 20,
        color: "111827",
        valign: "top",
        lineSpacingMultiple: 1.2,
      });

      // Accent callout chip
      s.addShape(pptx.ShapeType.roundRect, {
        x: 0.75,
        y: 6.55,
        w: 3.5,
        h: 0.55,
        fill: { color: secondary },
        line: { color: secondary },
        radius: 0.2,
      });
      s.addText("Ocean Professional", {
        x: 0.75,
        y: 6.63,
        w: 3.5,
        h: 0.45,
        fontFace: "Aptos",
        fontSize: 13,
        color: "111827",
        align: "center",
        bold: true,
      });

      addFooter(s, idx + 1, total);
    });

    // PptxGenJS provides ArrayBuffer in browser
    const arrayBuffer = await pptx.write("arraybuffer");
    return new Blob([arrayBuffer], {
      type: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    });
  }

  /**
   * Extracts a lightweight preview (titles + bullets) from generated PPTX.
   * We intentionally keep this fast and dependency-light by reading slide XML.
   *
   * @param {Blob} blob
   * @returns {Promise<SlidePreview[]>}
   */
  async function buildPreviewFromPptx(blob) {
    const buffer = await blob.arrayBuffer();
    const zip = await JSZip.loadAsync(buffer);

    const parser = new XMLParser({
      ignoreAttributes: false,
      attributeNamePrefix: "@_",
      removeNSPrefix: true, // easier to navigate
      preserveOrder: false,
      trimValues: true,
    });

    // slide files: ppt/slides/slide1.xml ...
    const slideFiles = Object.keys(zip.files)
      .filter((p) => /^ppt\/slides\/slide\d+\.xml$/.test(p))
      .sort((a, b) => {
        const ai = Number(a.match(/slide(\d+)\.xml/)?.[1] || 0);
        const bi = Number(b.match(/slide(\d+)\.xml/)?.[1] || 0);
        return ai - bi;
      });

    const previews = [];

    // Helper to recursively collect text nodes in typical DrawingML structure
    function collectText(node, acc) {
      if (!node || typeof node !== "object") return;

      if (typeof node === "string") return;

      // Common text run value appears as: a:t or t (after NS stripped)
      if (node.t && typeof node.t === "string") {
        acc.push(node.t);
      }

      for (const key of Object.keys(node)) {
        const val = node[key];
        if (Array.isArray(val)) {
          val.forEach((v) => collectText(v, acc));
        } else if (typeof val === "object") {
          collectText(val, acc);
        }
      }
    }

    for (let i = 0; i < slideFiles.length; i++) {
      const path = slideFiles[i];
      const xml = await zip.file(path).async("text");
      const json = parser.parse(xml);

      const texts = [];
      collectText(json, texts);

      // Heuristic: first non-empty string -> title, remaining -> bullets-ish
      const cleaned = texts.map((t) => String(t).replace(/\s+/g, " ").trim()).filter(Boolean);

      const title = cleaned[0] || `Slide ${i + 1}`;
      const rest = cleaned.slice(1);

      // Convert to bullet list; de-dup consecutive repeats
      const bullets = [];
      for (const t of rest) {
        if (!t) continue;
        if (bullets.length > 0 && bullets[bullets.length - 1] === t) continue;
        bullets.push(t);
        if (bullets.length >= 8) break;
      }

      previews.push({
        index: i,
        title,
        bullets,
      });
    }

    return previews.length
      ? previews
      : slides.map((s, i) => ({
          index: i,
          title: s.title || `Slide ${i + 1}`,
          bullets: parseBullets(s.bulletsRaw),
        }));
  }

  // PUBLIC_INTERFACE
  async function handleGenerate() {
    try {
      setStatus({ kind: "generating", message: "Generating PPTX and building preview…" });

      const filename = buildFilename();
      const blob = await generatePptxBlob();
      const previews = await buildPreviewFromPptx(blob);

      setPptxFilename(filename);
      setPptxBlob(blob);
      setPreviewSlides(previews);
      setActiveSlideIndex(0);

      setStatus({
        kind: "ready",
        message: "PPTX ready. Preview updated — you can download now.",
      });
    } catch (e) {
      // eslint-disable-next-line no-console
      console.error(e);
      setStatus({
        kind: "error",
        message:
          "Could not generate PPTX. Please reduce content and try again (details in console).",
      });
    }
  }

  // PUBLIC_INTERFACE
  function handleDownload() {
    if (!pptxBlob) return;
    saveAs(pptxBlob, pptxFilename);
  }

  function updateSlideField(index, patch) {
    setSlides((prev) => {
      const next = [...prev];
      next[index] = { ...next[index], ...patch };
      return next;
    });
  }

  function addSlide() {
    setSlideCount((n) => Math.min(n + 1, 20));
    // focus sidebar on the newly added slide a moment later
    window.setTimeout(() => {
      setActiveSlideIndex((idx) => Math.min(idx + 1, 19));
      leftPanelRef.current?.focus?.();
    }, 0);
  }

  function removeSlide(index) {
    if (slides.length <= 1) return;
    setSlides((prev) => prev.filter((_, i) => i !== index));
    setSlideCount((n) => Math.max(n - 1, 1));
    setActiveSlideIndex((idx) => Math.max(0, Math.min(idx, slides.length - 2)));
  }

  const effectivePreview = previewSlides.length
    ? previewSlides
    : slides.map((s, i) => ({
        index: i,
        title: s.title || `Slide ${i + 1}`,
        bullets: parseBullets(s.bulletsRaw),
      }));

  const currentPreview = effectivePreview[activeSlideIndex];

  return (
    <div className="OceanApp" style={{ background: theme.bg, color: theme.text }}>
      <a className="SkipLink" href="#mainPreview">
        Skip to preview
      </a>

      <header className="TopBar">
        <div className="TopBar__left">
          <div className="BrandMark" aria-hidden="true" />
          <div className="BrandText">
            <div className="BrandTitle">PPT Generator</div>
            <div className="BrandSub">Ocean Professional • in-browser • no backend</div>
          </div>
        </div>

        <div className="TopBar__right" role="group" aria-label="Primary actions">
          <button
            type="button"
            className="Btn Btn--primary"
            onClick={handleGenerate}
            disabled={status.kind === "generating"}
          >
            {status.kind === "generating" ? "Generating…" : "Generate PPTX"}
          </button>
          <button
            type="button"
            className="Btn Btn--secondary"
            onClick={handleDownload}
            disabled={!pptxBlob || status.kind === "generating"}
          >
            Download
          </button>
        </div>
      </header>

      <div className="Layout">
        <aside className="Sidebar" aria-label="Presentation controls">
          <div className="Card">
            <h2 className="CardTitle">Deck</h2>

            <label className="Field">
              <span className="FieldLabel">Title</span>
              <input
                type="text"
                className="Input"
                value={deckTitle}
                onChange={(e) => setDeckTitle(e.target.value)}
                placeholder="e.g., Q1 Business Review"
              />
            </label>

            <div className="FieldRow">
              <label className="Field" style={{ flex: 1 }}>
                <span className="FieldLabel">Author</span>
                <input
                  type="text"
                  className="Input"
                  value={author}
                  onChange={(e) => setAuthor(e.target.value)}
                  placeholder="e.g., Alex Doe"
                />
              </label>

              <label className="Field" style={{ width: 160 }}>
                <span className="FieldLabel">Accent</span>
                <select
                  className="Select"
                  value={accentMode}
                  onChange={(e) => setAccentMode(e.target.value)}
                >
                  <option value="gradient">Gradient</option>
                  <option value="solid">Solid</option>
                </select>
              </label>
            </div>

            <label className="Field">
              <span className="FieldLabel">
                Slides: <span className="Mono">{slideCount}</span>
              </span>
              <input
                type="range"
                min={1}
                max={12}
                value={slideCount}
                onChange={(e) => setSlideCount(Number(e.target.value))}
                className="Range"
              />
              <div className="Hint">Tip: keep it under 12 for fast in-browser preview parsing.</div>
            </label>

            <div className="Status" role="status" aria-live="polite" data-kind={status.kind}>
              <span className="StatusDot" aria-hidden="true" />
              <span>{status.message}</span>
            </div>
          </div>

          <div className="Card" tabIndex={-1} ref={leftPanelRef}>
            <div className="CardHeaderRow">
              <h2 className="CardTitle">Slides</h2>
              <button type="button" className="Btn Btn--ghost" onClick={addSlide}>
                + Add
              </button>
            </div>

            <div className="SlideList" role="listbox" aria-label="Slide list">
              {slides.map((s, idx) => (
                <button
                  key={idx}
                  type="button"
                  className={`SlideListItem ${idx === activeSlideIndex ? "is-active" : ""}`}
                  onClick={() => setActiveSlideIndex(idx)}
                  role="option"
                  aria-selected={idx === activeSlideIndex}
                >
                  <div className="SlideListItem__top">
                    <span className="Pill">{idx + 1}</span>
                    <span className="SlideListItem__title">
                      {s.title?.trim() ? s.title : `Slide ${idx + 1}`}
                    </span>
                  </div>
                  <div className="SlideListItem__meta">
                    {parseBullets(s.bulletsRaw).length} bullets
                  </div>
                </button>
              ))}
            </div>

            <div className="Divider" />

            <div className="Editor">
              <div className="EditorHeader">
                <h3 className="EditorTitle">Edit slide {activeSlideIndex + 1}</h3>
                <button
                  type="button"
                  className="Btn Btn--danger"
                  onClick={() => removeSlide(activeSlideIndex)}
                  disabled={slides.length <= 1}
                >
                  Remove
                </button>
              </div>

              <label className="Field">
                <span className="FieldLabel">Slide title</span>
                <input
                  type="text"
                  className="Input"
                  value={activeSlide?.title || ""}
                  onChange={(e) => updateSlideField(activeSlideIndex, { title: e.target.value })}
                />
              </label>

              <label className="Field">
                <span className="FieldLabel">Bullets (one per line)</span>
                <textarea
                  className="Textarea"
                  rows={7}
                  value={activeSlide?.bulletsRaw || ""}
                  onChange={(e) =>
                    updateSlideField(activeSlideIndex, { bulletsRaw: e.target.value })
                  }
                />
                <div className="Hint">
                  Lines become bullet points. Keep each bullet short for best slide layout.
                </div>
              </label>
            </div>
          </div>
        </aside>

        <main id="mainPreview" className="Main" aria-label="Preview">
          <div className="PreviewHeader">
            <div>
              <h2 className="PreviewTitle">Preview</h2>
              <div className="PreviewSub">
                {pptxBlob ? (
                  <>
                    Showing preview from generated PPTX •{" "}
                    <span className="Mono">{pptxFilename}</span>
                  </>
                ) : (
                  <>Preview updates instantly; generate to confirm final PPTX.</>
                )}
              </div>
            </div>

            <div className="PreviewNav" role="group" aria-label="Preview navigation">
              <button
                type="button"
                className="Btn Btn--ghost"
                onClick={() => setActiveSlideIndex((i) => Math.max(i - 1, 0))}
                disabled={activeSlideIndex <= 0}
              >
                ← Prev
              </button>
              <button
                type="button"
                className="Btn Btn--ghost"
                onClick={() =>
                  setActiveSlideIndex((i) => Math.min(i + 1, slides.length - 1))
                }
                disabled={activeSlideIndex >= slides.length - 1}
              >
                Next →
              </button>
            </div>
          </div>

          <section className="SlideStage" aria-label="Slide canvas">
            <div className="SlideCanvas" role="img" aria-label={`Slide ${activeSlideIndex + 1}`}>
              <div
                className={`SlideAccent ${accentMode === "gradient" ? "is-gradient" : "is-solid"}`}
                style={{
                  background:
                    accentMode === "gradient"
                      ? `linear-gradient(90deg, ${theme.primary} 0%, ${theme.primary} 65%, rgba(245, 158, 11, 0.85) 100%)`
                      : theme.primary,
                }}
              >
                {accentMode === "gradient" && <div className="SlideAccentGlow" aria-hidden="true" />}
              </div>

              <div className="SlideBody">
                <h3 className="SlideTitleText">{currentPreview?.title || "Slide"}</h3>
                <div className="SlideDivider" />
                <ul className="SlideBullets">
                  {(currentPreview?.bullets?.length ? currentPreview.bullets : ["Add bullets to the slide."]).map(
                    (b, i) => (
                      <li key={i}>{b}</li>
                    )
                  )}
                </ul>
              </div>

              <div className="SlideFooter">
                <span className="SlideFooter__left">{author || " "}</span>
                <span className="SlideFooter__right">
                  {activeSlideIndex + 1}/{slides.length}
                </span>
              </div>
            </div>

            <div className="PreviewRail" aria-label="Slide thumbnails">
              {effectivePreview.map((p, idx) => (
                <button
                  key={idx}
                  type="button"
                  className={`Thumb ${idx === activeSlideIndex ? "is-active" : ""}`}
                  onClick={() => setActiveSlideIndex(idx)}
                  aria-label={`Go to slide ${idx + 1}`}
                  aria-current={idx === activeSlideIndex ? "page" : undefined}
                >
                  <div className="Thumb__num">{idx + 1}</div>
                  <div className="Thumb__title">{p.title}</div>
                </button>
              ))}
            </div>
          </section>

          <section className="HelpCard" aria-label="How it works">
            <h3 className="HelpTitle">How this works</h3>
            <ul className="HelpList">
              <li>All generation happens in your browser using PptxGenJS.</li>
              <li>
                The preview reads slide XML from the generated PPTX (fast, lightweight, and no server
                required).
              </li>
              <li>
                Download saves the generated file using the browser’s file APIs (no backend).
              </li>
            </ul>
          </section>
        </main>
      </div>
    </div>
  );
}

export default App;
