import { useState } from 'react';
import type { ClientInfo } from '../../types';

interface Props {
  info: ClientInfo;
  onChange: (info: ClientInfo) => void;
  onNext: () => void;
  onBack: () => void;
}

function Field({
  label, id, type = 'text', placeholder, value, onChange, required,
}: {
  label: string; id: string; type?: string; placeholder: string;
  value: string; onChange: (v: string) => void; required?: boolean;
}) {
  return (
    <div>
      <label className="label" htmlFor={id} style={{ display: 'block', marginBottom: 6 }}>
        {label}{required && ' *'}
      </label>
      <input
        id={id}
        className="input"
        type={type}
        placeholder={placeholder}
        value={value}
        onChange={e => onChange(e.target.value)}
        autoComplete={type === 'email' ? 'email' : type === 'tel' ? 'tel' : 'on'}
      />
    </div>
  );
}

export default function ClientForm({ info, onChange, onNext }: Props) {
  const [errors, setErrors] = useState<Partial<Record<keyof ClientInfo, string>>>({});

  function set(key: keyof ClientInfo) {
    return (val: string) => {
      onChange({ ...info, [key]: val });
      setErrors(prev => ({ ...prev, [key]: undefined }));
    };
  }

  function validate(): boolean {
    const errs: Partial<Record<keyof ClientInfo, string>> = {};
    if (!info.nom.trim())       errs.nom       = 'Requis';
    if (!info.prenom.trim())    errs.prenom    = 'Requis';
    if (!info.telephone.trim()) errs.telephone = 'Requis';
    /* Email est optionnel */
    if (info.email.trim() && !/^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(info.email.trim())) {
      errs.email = 'Email invalide (ex: marie@exemple.fr)';
    }
    setErrors(errs);
    return Object.keys(errs).length === 0;
  }

  function handleNext() {
    if (validate()) onNext();
  }

  // Deadline badge
  function deadlineBadge(val: string) {
    if (!val) return null;
    const today = new Date(); today.setHours(0, 0, 0, 0);
    const dead = new Date(val + 'T00:00:00');
    const diff = Math.round((dead.getTime() - today.getTime()) / (1000 * 60 * 60 * 24));
    const cls = diff < 0 ? 'urgent' : diff <= 3 ? 'soon' : 'ok';
    const text = diff < 0 ? 'Dépassé' : diff === 0 ? "Aujourd'hui" : `${diff} j`;
    const bg = cls === 'urgent' ? '#FF3B30' : cls === 'soon' ? '#FF9500' : '#34C759';
    return (
      <span style={{
        display: 'inline-block', background: bg, color: 'white',
        borderRadius: 8, padding: '2px 8px', fontSize: 12, fontWeight: 700, marginLeft: 8,
      }}>
        {text}
      </span>
    );
  }

  return (
    <div className="step-panel" style={{ gap: 16 }}>

      <div className="card">
        <div className="card-title">Identité</div>
        <div className="card-body" style={{ gap: 12 }}>
          <div className="grid-2">
            <div>
              <label className="label" style={{ display: 'block', marginBottom: 6 }}>Nom *</label>
              <input
                className={`input ${errors.nom ? 'input-error' : ''}`}
                type="text"
                placeholder="Dupont"
                value={info.nom}
                onChange={e => {
                  const v = e.target.value.replace(/\b\w/g, l => l.toUpperCase());
                  set('nom')(v);
                }}
                style={{ borderColor: errors.nom ? 'var(--red)' : '' }}
              />
            </div>
            <div>
              <label className="label" style={{ display: 'block', marginBottom: 6 }}>Prénom *</label>
              <input
                className="input"
                type="text"
                placeholder="Marie"
                value={info.prenom}
                onChange={e => {
                  const v = e.target.value.replace(/\b\w/g, l => l.toUpperCase());
                  set('prenom')(v);
                }}
                style={{ borderColor: errors.prenom ? 'var(--red)' : '' }}
              />
            </div>
          </div>

          <Field
            label="Téléphone" id="tel" type="tel"
            placeholder="+33 6 00 00 00 00"
            value={info.telephone} onChange={set('telephone')}
            required
          />
          {errors.telephone && (
            <div style={{ color: 'var(--red)', fontSize: 12, marginTop: -8 }}>{errors.telephone}</div>
          )}

          <Field
            label="Email" id="email" type="email"
            placeholder="marie@exemple.fr (optionnel)"
            value={info.email} onChange={set('email')}
          />
          {errors.email && (
            <div style={{ color: 'var(--red)', fontSize: 12, marginTop: -8 }}>{errors.email}</div>
          )}
        </div>
      </div>

      <div className="card">
        <div className="card-title">Livraison</div>
        <div className="card-body" style={{ gap: 12 }}>
          <div>
            <label className="label" style={{ display: 'block', marginBottom: 6 }}>
              Adresse de livraison
            </label>
            <textarea
              className="input"
              rows={3}
              placeholder="12 rue de la Paix, 75001 Paris"
              value={info.adresse}
              onChange={e => set('adresse')(e.target.value)}
              style={{ resize: 'vertical' }}
            />
          </div>

          <div>
            <label className="label" style={{ display: 'block', marginBottom: 6 }}>
              Date limite
              {info.deadline && deadlineBadge(info.deadline)}
            </label>
            <input
              className="input"
              type="date"
              value={info.deadline}
              onChange={e => set('deadline')(e.target.value)}
              min={new Date().toISOString().split('T')[0]}
            />
          </div>
        </div>
      </div>

      <div className="step-hint">
        Ces informations sont transmises à l'atelier avec votre commande.
      </div>

      <button className="btn btn-primary" onClick={handleNext}>
        Continuer vers le paiement →
      </button>
    </div>
  );
}
