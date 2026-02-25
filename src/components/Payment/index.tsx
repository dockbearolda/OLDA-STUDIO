import { useState } from 'react';
import { motion, AnimatePresence } from 'framer-motion';
import { useCart } from '../../context/CartContext';
import type { ClientInfo, PaymentStatus } from '../../types';

interface Props {
  clientInfo: ClientInfo;
  onBack: () => void;
}

// Generate order number like: Nom-Prenom-DDMMYY-HHMM
function generateOrderNo(info: ClientInfo): string {
  const now = new Date();
  const dd = String(now.getDate()).padStart(2, '0');
  const mm = String(now.getMonth() + 1).padStart(2, '0');
  const yy = String(now.getFullYear()).slice(-2);
  const hh = String(now.getHours()).padStart(2, '0');
  const mi = String(now.getMinutes()).padStart(2, '0');
  const cap = (s: string) => s ? s.toLowerCase().replace(/(^|\s)\S/g, c => c.toUpperCase()) : '';
  const nom    = cap(info.nom.trim()).replace(/\s+/g, '-');
  const prenom = cap(info.prenom.trim()).replace(/\s+/g, '-');
  const namePart = nom && prenom ? `${nom}-${prenom}` : nom || prenom || 'Client';
  return `${namePart}-${dd}${mm}${yy}-${hh}${mi}`;
}

export default function Payment({ clientInfo, onBack }: Props) {
  const { items, total, clearCart } = useCart();
  const [payStatus, setPayStatus]   = useState<PaymentStatus>('unpaid');
  const [loading, setLoading]       = useState(false);
  const [success, setSuccess]       = useState(false);
  const [error, setError]           = useState('');

  // ── Build DASHOLDA payload ─────────────────────────────
  async function submit() {
    if (items.length === 0) {
      setError('Votre panier est vide. Ajoutez au moins un article.');
      return;
    }
    setLoading(true);
    setError('');

    try {
      // Fetch config (token + URL)
      const cfgRes = await fetch('/api/config');
      const cfg = await cfgRes.json() as { dashboard_token: string };
      const token = cfg.dashboard_token;

      const orderNo = generateOrderNo(clientInfo);

      // Build a multi-item order payload
      const orderPayload = {
        commande:     orderNo,
        nom:          clientInfo.nom,
        prenom:       clientInfo.prenom,
        email:        clientInfo.email,
        telephone:    clientInfo.telephone,
        adresse:      clientInfo.adresse,
        deadline:     clientInfo.deadline,
        famille:      items[0]?.famille ?? 'textile',
        // Items list for multi-article support
        items: items.map(item => ({
          collection:       item.collection,
          reference:        item.reference,
          couleurTshirt:    item.couleur.n,
          taille:           item.taille,
          note:             item.note,
          logoAvant:        item.logoAvant.name ?? '',
          logoArriere:      item.logoArriere.name ?? '',
          prix: {
            tshirt:          item.prix.tshirt,
            personnalisation: item.prix.personnalisation,
            total:            item.prix.total,
          },
        })),
        prix: {
          total: String(total),
          tshirt: String(items.reduce((s, i) => s + i.prix.tshirt, 0)),
          personnalisation: String(items.reduce((s, i) => s + i.prix.personnalisation, 0)),
        },
        paiement: { statut: payStatus === 'paid' ? 'OUI' : 'NON' },
        _reçuLe: new Date().toISOString(),
      };

      const res = await fetch('/api/webhook/forward-to-dasholda', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Authorization': `Bearer ${token}`,
        },
        body: JSON.stringify(orderPayload),
      });

      if (!res.ok) {
        const err = await res.json().catch(() => ({})) as { error?: string };
        throw new Error(err.error ?? `Erreur HTTP ${res.status}`);
      }

      // Success!
      clearCart();
      setSuccess(true);
    } catch (e) {
      setError(e instanceof Error ? e.message : 'Erreur lors de l\'envoi.');
    } finally {
      setLoading(false);
    }
  }

  // ── Success screen ─────────────────────────────────────
  if (success) {
    return (
      <div className="step-panel">
        <div className="card">
          <div className="card-body" style={{ gap: 0 }}>
            <div className="success-screen">
              <motion.div
                className="success-icon"
                initial={{ scale: 0 }}
                animate={{ scale: 1 }}
                transition={{ type: 'spring', stiffness: 400, damping: 20, delay: 0.1 }}
              >
                ✓
              </motion.div>
              <motion.div
                className="success-title"
                initial={{ opacity: 0, y: 12 }}
                animate={{ opacity: 1, y: 0 }}
                transition={{ delay: 0.25 }}
              >
                Commande envoyée !
              </motion.div>
              <motion.div
                className="success-sub"
                initial={{ opacity: 0 }}
                animate={{ opacity: 1 }}
                transition={{ delay: 0.4 }}
              >
                Votre commande a été transmise au Dashboard de l'atelier.
                Mélina vous contactera dès que possible.
              </motion.div>
              <motion.button
                className="btn btn-primary"
                style={{ marginTop: 8 }}
                initial={{ opacity: 0 }}
                animate={{ opacity: 1 }}
                transition={{ delay: 0.55 }}
                onClick={() => window.location.reload()}
              >
                Nouvelle commande
              </motion.button>
            </div>
          </div>
        </div>
      </div>
    );
  }

  const ht  = Math.round(total / 1.2 * 100) / 100;
  const tva = Math.round((total - ht) * 100) / 100;

  return (
    <div className="step-panel" style={{ gap: 16 }}>

      {/* ── ORDER SUMMARY ─────────────────────────── */}
      <div className="card">
        <div className="card-title">Récapitulatif de la commande</div>
        <div className="card-body" style={{ gap: 0 }}>
          {items.length === 0 ? (
            <div style={{ padding: '12px 0', color: 'var(--text-3)', textAlign: 'center', fontSize: 14 }}>
              Panier vide — retournez au Studio pour ajouter des articles.
            </div>
          ) : (
            items.map(item => (
              <div key={item.id} className="summary-item">
                <div
                  className="summary-color"
                  style={{ background: item.couleur.h }}
                  title={item.couleur.n}
                />
                <div className="summary-info">
                  <div className="summary-title">{item.couleur.n} · {item.taille}</div>
                  <div className="summary-sub">
                    {[item.collection, item.reference].filter(Boolean).join(' ')}
                    {item.logoAvant.id && ' · Logo ✓'}
                    {item.note && ` · ${item.note.slice(0, 40)}`}
                  </div>
                </div>
                <div className="summary-price">{item.prix.total} €</div>
              </div>
            ))
          )}
        </div>
      </div>

      {/* ── CLIENT INFO RECAP ──────────────────────── */}
      <div className="card">
        <div className="card-title">Client</div>
        <div className="card-body" style={{ gap: 6 }}>
          <div className="flex-between">
            <span className="text-sm text-2">Nom</span>
            <span className="text-sm fw-600">{clientInfo.nom} {clientInfo.prenom}</span>
          </div>
          <div className="flex-between">
            <span className="text-sm text-2">Téléphone</span>
            <span className="text-sm fw-600">{clientInfo.telephone || '—'}</span>
          </div>
          <div className="flex-between">
            <span className="text-sm text-2">Email</span>
            <span className="text-sm fw-600">{clientInfo.email || '—'}</span>
          </div>
          {clientInfo.adresse && (
            <div className="flex-between" style={{ alignItems: 'flex-start' }}>
              <span className="text-sm text-2">Adresse</span>
              <span className="text-sm fw-600" style={{ textAlign: 'right', maxWidth: '60%' }}>
                {clientInfo.adresse}
              </span>
            </div>
          )}
          {clientInfo.deadline && (
            <div className="flex-between">
              <span className="text-sm text-2">Deadline</span>
              <span className="text-sm fw-600">{clientInfo.deadline}</span>
            </div>
          )}
        </div>
      </div>

      {/* ── TOTAL ─────────────────────────────────── */}
      <div className="total-bubble">
        <div className="total-row">
          <span>Montant HT</span>
          <span>{ht.toFixed(2)} €</span>
        </div>
        <div className="total-row">
          <span>TVA (20%)</span>
          <span>{tva.toFixed(2)} €</span>
        </div>
        <div className="total-main">
          <span className="total-label">Total TTC</span>
          <span className="total-amount">{total} €</span>
        </div>
      </div>

      {/* ── PAYMENT STATUS ────────────────────────── */}
      <div className="card">
        <div className="card-title">Statut du paiement</div>
        <div className="card-body" style={{ gap: 8 }}>
          <button
            className={`pay-opt ${payStatus === 'paid' ? 'paid' : ''}`}
            onClick={() => setPayStatus('paid')}
          >
            <span>💳 Payé</span>
          </button>
          <button
            className={`pay-opt ${payStatus === 'unpaid' ? 'unpaid' : ''}`}
            onClick={() => setPayStatus('unpaid')}
          >
            <span>⏳ À régler</span>
          </button>
        </div>
      </div>

      {/* ── ERROR ─────────────────────────────────── */}
      <AnimatePresence>
        {error && (
          <motion.div
            initial={{ opacity: 0, y: -8 }}
            animate={{ opacity: 1, y: 0 }}
            exit={{ opacity: 0 }}
            style={{
              background: 'rgba(255,59,48,0.08)',
              border: '1px solid rgba(255,59,48,0.25)',
              borderRadius: 10,
              padding: '10px 14px',
              fontSize: 14,
              color: 'var(--red)',
            }}
          >
            {error}
          </motion.div>
        )}
      </AnimatePresence>

      {/* ── SUBMIT ────────────────────────────────── */}
      <button
        className="btn btn-primary"
        onClick={submit}
        disabled={loading || items.length === 0}
        style={{ opacity: items.length === 0 ? 0.5 : 1 }}
      >
        {loading ? (
          <><span className="btn-spinner" /> Envoi en cours…</>
        ) : (
          '✓ Envoyer à l\'Atelier'
        )}
      </button>

      <button className="btn btn-ghost" onClick={onBack} disabled={loading}>
        ← Modifier les informations
      </button>
    </div>
  );
}
