import {
  useState, useEffect, useRef, useCallback, type PointerEvent,
} from 'react';
import { motion, AnimatePresence } from 'framer-motion';
import { useCart } from '../../context/CartContext';
import {
  COLORS, LOGO_COLORS, REFERENCES, COLLECTIONS, SIZES,
  PRICE_TSHIRT, PRICE_PERSO, LOGO_SECTIONS,
  applyFabricColor, isLightColor,
} from '../../data/products';
import type { CartItem, Color, LogoPlacement } from '../../types';

// ── Default logo placement values ──────────────────────
const DEFAULT_LOGO: LogoPlacement = {
  id: null, type: null, svg: null, url: null, name: null,
  x: 66.5, y: 28, w: 19, color: '#1A1A1A',
};
const DEFAULT_LOGO_BACK: LogoPlacement = {
  id: null, type: null, svg: null, url: null, name: null,
  x: 47.9, y: 31, w: 30, color: '#1A1A1A',
};

// ── Build refs for a collection prefix ─────────────────
function buildRefs(prefix: string) {
  return Array.from({ length: 10 }, (_, i) => {
    const code = `${prefix.toUpperCase()[0]}-${String(i + 1).padStart(3, '0')}`;
    const ref = REFERENCES[code];
    const label = ref ? `${code}  ·  ${ref.fournisseur}` : code;
    return { value: code, label };
  });
}

// ── Generate a unique cart item id ─────────────────────
function uid(): string {
  return Math.random().toString(36).slice(2) + Date.now().toString(36);
}

interface Props {
  onNext: () => void;
}

export default function Studio({ onNext }: Props) {
  const { addItem } = useCart();

  // ── Product config state ───────────────────────────────
  const [famille, setFamille]     = useState<'textile' | 'mug'>('textile');
  const [collection, setCollection] = useState('');
  const [reference, setReference] = useState('');
  const [taille, setTaille]       = useState('M');
  const [color, setColor]         = useState<Color>(COLORS[0]);
  const [note, setNote]           = useState('');
  const [prixTshirt, setPrixTshirt] = useState(25);
  const [prixPerso, setPrixPerso]   = useState(0);

  // ── T-shirt visual state ───────────────────────────────
  const [side, setSide]           = useState<'front' | 'back'>('front');
  const [svgFront, setSvgFront]   = useState('');
  const [svgBack, setSvgBack]     = useState('');
  const [logoAvant, setLogoAvant]    = useState<LogoPlacement>({ ...DEFAULT_LOGO });
  const [logoArriere, setLogoArriere] = useState<LogoPlacement>({ ...DEFAULT_LOGO_BACK });

  // ── Logo sheet ─────────────────────────────────────────
  const [showSheet, setShowSheet] = useState(false);

  // ── Drag state ─────────────────────────────────────────
  const [dragging, setDragging] = useState(false);
  const stageRef = useRef<HTMLDivElement>(null);
  const dragOffsetRef = useRef({ dx: 0, dy: 0 });

  // ── Fetch SVGs ─────────────────────────────────────────
  useEffect(() => {
    fetch('/tshirt-front.svg').then(r => r.text()).then(setSvgFront).catch(() => {});
    fetch('/tshirt-back.svg').then(r => r.text()).then(setSvgBack).catch(() => {});
  }, []);

  // ── Auto-price from reference ──────────────────────────
  useEffect(() => {
    const ref = REFERENCES[reference];
    if (ref) setPrixTshirt(ref.prix);
  }, [reference]);

  // ── Computed values ────────────────────────────────────
  const total     = prixTshirt + prixPerso;
  const refOpts   = collection ? buildRefs(collection) : [];
  const currentLogo = side === 'front' ? logoAvant : logoArriere;
  const setCurrentLogo = side === 'front' ? setLogoAvant : setLogoArriere;

  // ── Render colored SVG ─────────────────────────────────
  function renderSvg(raw: string): string {
    return applyFabricColor(raw, color.h);
  }

  // ── Logo drag handlers ─────────────────────────────────
  const onLogoPointerDown = useCallback((e: PointerEvent<HTMLDivElement>) => {
    if (!currentLogo.id && !currentLogo.url) return;
    e.currentTarget.setPointerCapture(e.pointerId);
    setDragging(true);
    const stage = stageRef.current;
    if (!stage) return;
    const rect = stage.getBoundingClientRect();
    const logoX = (currentLogo.x / 100) * rect.width;
    const logoY = (currentLogo.y / 100) * rect.height;
    dragOffsetRef.current = {
      dx: e.clientX - rect.left - logoX,
      dy: e.clientY - rect.top  - logoY,
    };
  }, [currentLogo]);

  const onLogoPointerMove = useCallback((e: PointerEvent<HTMLDivElement>) => {
    if (!dragging) return;
    const stage = stageRef.current;
    if (!stage) return;
    const rect = stage.getBoundingClientRect();
    const rawX = e.clientX - rect.left - dragOffsetRef.current.dx;
    const rawY = e.clientY - rect.top  - dragOffsetRef.current.dy;
    const newX = Math.min(100, Math.max(0, (rawX / rect.width)  * 100));
    const newY = Math.min(100, Math.max(0, (rawY / rect.height) * 100));
    setCurrentLogo(prev => ({ ...prev, x: newX, y: newY }));
  }, [dragging, setCurrentLogo]);

  const onLogoPointerUp = useCallback(() => {
    setDragging(false);
  }, []);

  // ── Pick logo from sheet ───────────────────────────────
  function pickAtelier(logoId: string, svg: string, name: string) {
    const pos = side === 'front'
      ? { x: 66.5, y: 28, w: 19 }
      : { x: 47.9, y: 31, w: 30 };
    setCurrentLogo(prev => ({
      ...prev, id: logoId, type: 'atelier', svg, url: null, name, ...pos,
    }));
    setShowSheet(false);
  }

  function pickUpload(e: React.ChangeEvent<HTMLInputElement>) {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = ev => {
      const url = ev.target?.result as string;
      const pos = side === 'front'
        ? { x: 66.5, y: 28, w: 19 }
        : { x: 47.9, y: 31, w: 30 };
      setCurrentLogo(prev => ({
        ...prev, id: 'upload', type: 'upload', svg: null, url, name: file.name, ...pos,
      }));
      setShowSheet(false);
    };
    reader.readAsDataURL(file);
    e.target.value = '';
  }

  function removeLogo() {
    setCurrentLogo(side === 'front'
      ? { ...DEFAULT_LOGO }
      : { ...DEFAULT_LOGO_BACK }
    );
  }

  // ── Add to cart ────────────────────────────────────────
  function addToCart() {
    if (!collection || !reference || !taille) {
      alert('Veuillez sélectionner une collection, une référence et une taille.');
      return;
    }
    const item: CartItem = {
      id: uid(),
      famille,
      collection,
      reference,
      couleur: color,
      taille,
      logoAvant,
      logoArriere,
      note,
      prix: { tshirt: prixTshirt, personnalisation: prixPerso, total },
      addedAt: new Date().toISOString(),
    };
    addItem(item);

    // Reset visual state for next item
    setNote('');
    setReference('');
    setLogoAvant({ ...DEFAULT_LOGO });
    setLogoArriere({ ...DEFAULT_LOGO_BACK });
  }

  // ── Render logo overlay on SVG ─────────────────────────
  function renderLogoOverlay(logo: LogoPlacement, isActive: boolean) {
    if (!isActive) return null;
    if (!logo.id && !logo.url) {
      // Show tap pip
      const pipClass = side === 'front' ? 'logo-tap-front' : 'logo-tap-back';
      return (
        <div
          className={pipClass}
          onClick={() => setShowSheet(true)}
          title="Ajouter un logo"
        >
          <div className="logo-tap-pip">＋</div>
        </div>
      );
    }

    return (
      <div
        className="logo-zone"
        style={{
          left: `${logo.x}%`,
          top:  `${logo.y}%`,
          width: `${logo.w}%`,
          cursor: dragging ? 'grabbing' : 'grab',
        }}
        onPointerDown={onLogoPointerDown}
        onPointerMove={onLogoPointerMove}
        onPointerUp={onLogoPointerUp}
        onPointerCancel={onLogoPointerUp}
      >
        <div className="logo-inner">
          {logo.type === 'atelier' && logo.svg ? (
            <div
              style={{ color: logo.color, width: '100%', height: '100%', display: 'flex', alignItems: 'center', justifyContent: 'center' }}
              dangerouslySetInnerHTML={{ __html: logo.svg }}
            />
          ) : logo.url ? (
            <img src={logo.url} alt={logo.name ?? 'logo'} />
          ) : null}
        </div>
        <div className="logo-frame" />
      </div>
    );
  }

  return (
    <div className="step-panel" style={{ gap: 12 }}>

      {/* ── FAMILLE ───────────────────────────────── */}
      <div className="card">
        <div className="card-title">Type de produit</div>
        <div className="card-body">
          <div className="family-grid">
            {([['textile', '👕', 'T-Shirt', 'Personnalisé DTF'], ['mug', '☕', 'Mug', 'Sublimation']] as const).map(
              ([fam, icon, label, sub]) => (
                <button
                  key={fam}
                  className={`family-card ${famille === fam ? 'selected' : ''}`}
                  onClick={() => setFamille(fam)}
                >
                  <div className="family-card-icon">{icon}</div>
                  <div className="family-card-name">{label}</div>
                  <div className="family-card-sub">{sub}</div>
                </button>
              )
            )}
          </div>
        </div>
      </div>

      {/* ── T-SHIRT STUDIO ────────────────────────── */}
      {famille === 'textile' && (
        <>
          {/* T-shirt visual */}
          <div className="card">
            <div className="card-title">Aperçu</div>
            <div className="stage">
              <div className="shirts-grid">
                {(['front', 'back'] as const).map(s => {
                  const raw = s === 'front' ? svgFront : svgBack;
                  const logo = s === 'front' ? logoAvant : logoArriere;
                  const active = side === s;
                  return (
                    <div
                      key={s}
                      className={`shirt-col ${active ? 'active' : ''}`}
                      onClick={() => setSide(s)}
                    >
                      <div className="shirt-side-label">
                        {s === 'front' ? 'Avant' : 'Arrière'}
                      </div>
                      <div
                        className="svg-view"
                        ref={active ? stageRef : undefined}
                        style={{ position: 'relative' }}
                      >
                        {raw ? (
                          <div
                            className="svg-box"
                            dangerouslySetInnerHTML={{ __html: renderSvg(raw) }}
                          />
                        ) : (
                          <div style={{ width: '100%', height: '100%', background: 'rgba(0,0,0,0.04)', borderRadius: 8 }} />
                        )}
                        {renderLogoOverlay(logo, active)}
                      </div>
                    </div>
                  );
                })}
              </div>

              {/* Logo controls float bar */}
              <AnimatePresence>
                {(currentLogo.id || currentLogo.url) && (
                  <motion.div
                    className="logo-float-bar"
                    initial={{ opacity: 0, y: 8 }}
                    animate={{ opacity: 1, y: 0 }}
                    exit={{ opacity: 0, y: 8 }}
                    transition={{ duration: 0.2 }}
                  >
                    <span className="fbar-label">
                      {currentLogo.name ?? 'Logo'} · {side === 'front' ? 'Avant' : 'Arrière'}
                    </span>
                    <button
                      className="fbar-btn fbar-change"
                      onClick={() => setShowSheet(true)}
                    >
                      Changer
                    </button>
                    <button className="fbar-btn fbar-delete" onClick={removeLogo}>
                      Suppr.
                    </button>
                  </motion.div>
                )}
              </AnimatePresence>

              {/* Logo scale + color (when logo present) */}
              <AnimatePresence>
                {(currentLogo.id || currentLogo.url) && (
                  <motion.div
                    style={{ padding: '0 12px 10px' }}
                    initial={{ opacity: 0 }}
                    animate={{ opacity: 1 }}
                    exit={{ opacity: 0 }}
                  >
                    <div className="logo-scale-row">
                      <span className="text-xs text-3">Taille</span>
                      <input
                        type="range"
                        min="8"
                        max="70"
                        value={currentLogo.w}
                        onChange={e => setCurrentLogo(prev => ({ ...prev, w: +e.target.value }))}
                        style={{ flex: 1, accentColor: 'var(--accent)' }}
                      />
                      <span className="text-xs text-3">{Math.round(currentLogo.w)}%</span>
                    </div>
                    {currentLogo.type === 'atelier' && (
                      <div style={{ marginTop: 8 }}>
                        <div className="text-xs text-3" style={{ marginBottom: 6 }}>Couleur logo</div>
                        <div className="logo-colors">
                          {LOGO_COLORS.map(lc => (
                            <button
                              key={lc.h}
                              className={`logo-color-swatch ${currentLogo.color === lc.h ? 'selected' : ''}`}
                              style={{
                                background: lc.h,
                                boxShadow: lc.h === '#FFFFFF' ? 'inset 0 0 0 1px rgba(0,0,0,0.12)' : undefined,
                              }}
                              title={lc.n}
                              onClick={() => setCurrentLogo(prev => ({ ...prev, color: lc.h }))}
                            />
                          ))}
                        </div>
                      </div>
                    )}
                  </motion.div>
                )}
              </AnimatePresence>

              {/* Add logo tap area when no logo yet (not currently on the side) */}
              {!currentLogo.id && !currentLogo.url && (
                <div style={{ padding: '0 12px 10px' }}>
                  <button
                    className="btn btn-ghost btn-sm"
                    style={{ width: 'auto', paddingLeft: 14, paddingRight: 14 }}
                    onClick={() => setShowSheet(true)}
                  >
                    + Ajouter un logo ({side === 'front' ? 'Avant' : 'Arrière'})
                  </button>
                </div>
              )}
            </div>
          </div>

          {/* Color picker */}
          <div className="card">
            <div className="card-title">Couleur du t-shirt — {color.n}</div>
            <div className="card-body">
              <div className="swatches-grid">
                {COLORS.map(c => (
                  <button
                    key={c.h}
                    className={`swatch ${color.h === c.h ? 'selected' : ''}`}
                    style={{ background: c.h }}
                    data-light={isLightColor(c.h) ? '1' : '0'}
                    title={c.n}
                    onClick={() => setColor(c)}
                  />
                ))}
              </div>
            </div>
          </div>
        </>
      )}

      {/* ── PRODUCT DETAILS ───────────────────────── */}
      <div className="card">
        <div className="card-title">Référence produit</div>
        <div className="card-body">
          {/* Collection */}
          <div>
            <div className="label" style={{ marginBottom: 6 }}>Collection</div>
            <select
              className="native-select"
              value={collection}
              onChange={e => { setCollection(e.target.value); setReference(''); }}
            >
              <option value="">Choisir…</option>
              {COLLECTIONS.map(c => (
                <option key={c} value={c}>{c}</option>
              ))}
            </select>
          </div>

          {/* Ref + Size */}
          <div className="grid-2">
            <div>
              <div className="label" style={{ marginBottom: 6 }}>Référence</div>
              <select
                className="native-select"
                value={reference}
                onChange={e => setReference(e.target.value)}
                disabled={!collection}
              >
                <option value="">—</option>
                {refOpts.map(r => (
                  <option key={r.value} value={r.value}>{r.label}</option>
                ))}
              </select>
            </div>
            <div>
              <div className="label" style={{ marginBottom: 6 }}>Taille</div>
              <select
                className="native-select"
                value={taille}
                onChange={e => setTaille(e.target.value)}
              >
                {SIZES.map(s => (
                  <option key={s} value={s}>{s}</option>
                ))}
              </select>
            </div>
          </div>

          {/* Notes */}
          <div>
            <div className="label" style={{ marginBottom: 6 }}>Notes / instructions</div>
            <textarea
              className="input"
              rows={2}
              placeholder="Ex: logo sur la manche, couleur spécifique…"
              value={note}
              onChange={e => setNote(e.target.value)}
              style={{ resize: 'vertical', minHeight: 64 }}
            />
          </div>
        </div>
      </div>

      {/* ── PRICE ─────────────────────────────────── */}
      <div className="card">
        <div className="card-title">Tarification</div>
        <div className="card-body">
          <div className="grid-2">
            <div>
              <div className="label" style={{ marginBottom: 6 }}>Support</div>
              <select
                className="native-select"
                value={prixTshirt}
                onChange={e => setPrixTshirt(+e.target.value)}
              >
                {PRICE_TSHIRT.map(p => (
                  <option key={p} value={p}>{p} €</option>
                ))}
              </select>
            </div>
            <div>
              <div className="label" style={{ marginBottom: 6 }}>Personnalisation</div>
              <select
                className="native-select"
                value={prixPerso}
                onChange={e => setPrixPerso(+e.target.value)}
              >
                {PRICE_PERSO.map(p => (
                  <option key={p} value={p}>{p} €</option>
                ))}
              </select>
            </div>
          </div>

          <div className="total-bubble">
            <div className="total-main">
              <span className="total-label">Total article</span>
              <span className="total-amount">{total} €</span>
            </div>
          </div>
        </div>
      </div>

      {/* ── CTA ───────────────────────────────────── */}
      <button className="btn btn-dark" onClick={addToCart}>
        + Ajouter au panier
      </button>
      <button className="btn btn-primary" onClick={onNext}>
        Informations client →
      </button>

      {/* ── LOGO SELECTOR BOTTOM SHEET ────────────── */}
      <AnimatePresence>
        {showSheet && (
          <LogoSheet
            onClose={() => setShowSheet(false)}
            onPickAtelier={pickAtelier}
            onPickUpload={pickUpload}
            currentId={currentLogo.id}
          />
        )}
      </AnimatePresence>
    </div>
  );
}

// ── Logo Sheet ──────────────────────────────────────────
interface SheetProps {
  onClose: () => void;
  onPickAtelier: (id: string, svg: string, name: string) => void;
  onPickUpload: (e: React.ChangeEvent<HTMLInputElement>) => void;
  currentId: string | null;
}

function LogoSheet({ onClose, onPickAtelier, onPickUpload, currentId }: SheetProps) {
  const inputRef = useRef<HTMLInputElement>(null);
  return (
    <motion.div
      className="sheet-overlay"
      initial={{ opacity: 0 }}
      animate={{ opacity: 1 }}
      exit={{ opacity: 0 }}
      transition={{ duration: 0.2 }}
      onClick={onClose}
    >
      <motion.div
        className="sheet"
        initial={{ y: '100%' }}
        animate={{ y: 0 }}
        exit={{ y: '100%' }}
        transition={{ type: 'spring', stiffness: 400, damping: 40 }}
        onClick={e => e.stopPropagation()}
      >
        <div className="sheet-handle" />
        <div className="sheet-header">
          <span className="sheet-title">Choisir un logo</span>
          <button className="sheet-close" onClick={onClose}>×</button>
        </div>
        <div className="sheet-scroll">
          {LOGO_SECTIONS.map(section => (
            <div key={section.key}>
              <div className="sheet-section-title">{section.label}</div>
              <div className="logo-grid" style={{ marginBottom: 20 }}>
                {section.logos.map(logo => (
                  <button
                    key={logo.id}
                    className={`logo-item ${currentId === logo.id ? 'selected' : ''}`}
                    onClick={() => logo.s && onPickAtelier(logo.id, logo.s, `${section.label} · ${logo.n}`)}
                  >
                    <div className="logo-item-preview">
                      {logo.s && (
                        <div
                          style={{ color: '#1A1A1A', width: 52, height: 52, display: 'flex', alignItems: 'center', justifyContent: 'center' }}
                          dangerouslySetInnerHTML={{ __html: logo.s }}
                        />
                      )}
                    </div>
                    <span className="logo-item-name">{logo.n}</span>
                  </button>
                ))}
              </div>
            </div>
          ))}

          {/* Upload */}
          <div className="sheet-section-title">Votre logo</div>
          <div
            className="upload-row"
            onClick={() => inputRef.current?.click()}
          >
            <span style={{ fontSize: 22 }}>📁</span>
            <div>
              <div className="fw-600" style={{ fontSize: 14 }}>Importer un fichier</div>
              <div className="text-xs text-3">PNG, JPEG, SVG, WebP</div>
            </div>
          </div>
          <input
            ref={inputRef}
            type="file"
            accept="image/png,image/jpeg,image/svg+xml,image/webp"
            style={{ display: 'none' }}
            onChange={onPickUpload}
          />

          <div style={{ height: 32 }} />
        </div>
      </motion.div>
    </motion.div>
  );
}
