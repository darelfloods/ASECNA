import React, { useState, useEffect, useCallback } from 'react';
import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';

const API = 'http://localhost:3002/api';

// ── Types ─────────────────────────────────────────────────────────────────────

interface Fiche {
  id: string;
  numero_fiche: string;
  assistant: string;
  numero_vol: string;
  type_vol: 'regulier' | 'irregulier';
  date_saisie: string;
  site: string;
  compagnie_assistee: string;
  immatricule_aeronef: string;
  vol_national: boolean;
  vol_regional: boolean;
  vol_international: boolean;
  banques_depart: number[];
  nombre_banques_depart: number;
  heure_ouverture_depart: string;
  date_ouverture_depart: string;
  heure_cloture_depart: string;
  date_cloture_depart: string;
  pax_prevu_depart: number;
  banques_arrivee: number[];
  nombre_banques_arrivee: number;
  heure_ouverture_arrivee: string;
  date_ouverture_arrivee: string;
  heure_cloture_arrivee: string;
  date_cloture_arrivee: string;
  duree_comptoirs_minutes: number;
  pax_arrives: number;
  pax_departs: number;
  pax_transit: number;
  duree_heures_decimal: number;
  statut: 'saisie' | 'facturee';
}

interface Facture {
  id: string;
  numero_facture: string;
  date_facture: string;
  compagnie: string;
  adresse_compagnie?: string;
  ville_compagnie?: string;
  site: string;
  serie_bandes: string;
  periode_debut: string;
  periode_fin: string;
  fiches_ids: string[];
  nombre_heures: number;
  tarif_horaire: number;
  total_heures: number;
  nombre_annonces: number;
  tarif_annonce: number;
  total_annonces: number;
  montant_ht: number;
  taxes: number;
  acompte: number;
  solde: number;
  total_pax: number;
  montant_en_lettres: string;
  statut: 'brouillon' | 'emise' | 'payee';
}

// ── Constants ─────────────────────────────────────────────────────────────────

const SITES = ['LIBREVILLE', 'PORT-GENTIL', 'OYEM', 'BITAM', 'MVENGUE', 'FRANCEVILLE', 'LAMBARENE', 'MOANDA', 'MAKOKOU'];
const COMPAGNIES = ['HELI-UNION', 'AFRIC AVIATION', 'AFRIJET', 'EAGLE', 'ART', 'COREX INTERNATIONALE', 'BUSINESS BUREAU', 'LA NATIONALE', 'CRONOS AIR'];

// ── Utilities ─────────────────────────────────────────────────────────────────

function numberToWords(n: number): string {
  if (n === 0) return 'ZÉRO (0) Francs.CFA';

  const unites = ['', 'UN', 'DEUX', 'TROIS', 'QUATRE', 'CINQ', 'SIX', 'SEPT', 'HUIT', 'NEUF',
    'DIX', 'ONZE', 'DOUZE', 'TREIZE', 'QUATORZE', 'QUINZE', 'SEIZE', 'DIX-SEPT', 'DIX-HUIT', 'DIX-NEUF'];
  const dizaines = ['', 'DIX', 'VINGT', 'TRENTE', 'QUARANTE', 'CINQUANTE', 'SOIXANTE', 'SOIXANTE', 'QUATRE-VINGT', 'QUATRE-VINGT'];

  function convertHundreds(num: number, forMult = false): string {
    if (num === 0) return '';
    let res = '';
    const h = Math.floor(num / 100);
    const rest = num % 100;
    if (h > 0) {
      res += h === 1 ? 'CENT' : unites[h] + ' CENT';
      if (rest === 0 && h > 1 && !forMult) res += 'S';
      if (rest > 0) res += ' ';
    }
    if (rest > 0) {
      if (rest < 20) {
        res += unites[rest];
      } else {
        const tens = Math.floor(rest / 10);
        const units = rest % 10;
        if (tens === 7 || tens === 9) {
          res += dizaines[tens] + '-' + unites[10 + units];
        } else if (tens === 8) {
          res += 'QUATRE-VINGT';
          if (units > 0) res += '-' + unites[units];
          else if (!forMult) res += 'S';
        } else {
          res += dizaines[tens];
          if (units === 1 && tens > 1) res += '-ET-UN';
          else if (units > 0) res += '-' + unites[units];
        }
      }
    }
    return res;
  }

  const millions = Math.floor(n / 1_000_000);
  const thousands = Math.floor((n % 1_000_000) / 1_000);
  const remainder = n % 1_000;
  let res = '';

  if (millions > 0) {
    res += millions === 1 ? 'UN MILLION' : convertHundreds(millions) + ' MILLIONS';
    if (thousands > 0 || remainder > 0) res += ' ';
  }
  if (thousands > 0) {
    res += thousands === 1 ? 'MILLE' : convertHundreds(thousands, true) + ' MILLE';
    if (remainder > 0) res += ' ';
  }
  if (remainder > 0) res += convertHundreds(remainder);

  const formatted = n.toLocaleString('fr-FR').replace(/\u202f/g, '\u00a0');
  return `${res.trim()} (${formatted}) Francs.CFA`;
}

function calcDurationMinutes(openTime: string, openDate: string, closeTime: string, closeDate: string): number {
  if (!openTime || !closeTime) return 0;
  const [oh, om] = openTime.split(':').map(Number);
  const [ch, cm] = closeTime.split(':').map(Number);
  let diff = (ch * 60 + cm) - (oh * 60 + om);
  if (openDate && closeDate && openDate !== closeDate) {
    const open = new Date(openDate), close = new Date(closeDate);
    const days = Math.floor((close.getTime() - open.getTime()) / 86400000);
    diff += days * 1440;
  } else if (diff < 0) {
    diff += 1440;
  }
  return Math.max(0, diff);
}

function roundToHalf(minutes: number): number {
  if (minutes <= 0) return 0;
  if (minutes < 30) return 0.5;
  return Math.ceil((minutes / 60) * 2) / 2;
}

function minsToHHMM(m: number): string {
  const h = Math.floor(m / 60), min = m % 60;
  return `${h}h${String(min).padStart(2, '0')}`;
}

function fmt(n: number): string {
  return new Intl.NumberFormat('fr-FR').format(n).replace(/\u202f/g, '\u00a0');
}

function today(): string { return new Date().toISOString().split('T')[0]; }

function fmtDateFull(iso: string): string {
  if (!iso) return '';
  const [y, m, d] = iso.split('-');
  return `${d}/${m}/${y}`;
}
function fmtDateShort(iso: string): string {
  if (!iso) return '';
  const [, m, d] = iso.split('-');
  return `${d}/${m}`;
}
function buildPeriodeStr(debut: string, fin: string): string {
  if (!debut || !fin) return '';
  return `Du ${fmtDateShort(debut)} au ${fmtDateFull(fin)}`;
}

// ── API helpers ───────────────────────────────────────────────────────────────

async function apiFetch(method: string, path: string, body?: object) {
  const res = await fetch(`${API}${path}`, {
    method,
    headers: { 'Content-Type': 'application/json' },
    body: body ? JSON.stringify(body) : undefined,
  });
  return res.json();
}

// ── Toast ─────────────────────────────────────────────────────────────────────

function Toast({ msg, onClose }: { msg: string; onClose: () => void }) {
  useEffect(() => { const t = setTimeout(onClose, 3000); return () => clearTimeout(t); }, [onClose]);
  return (
    <div style={{
      position: 'fixed', bottom: 24, right: 24, background: '#1A2B4A', color: '#fff',
      padding: '12px 20px', borderRadius: 8, fontSize: 14, zIndex: 9999,
      boxShadow: '0 4px 16px rgba(0,0,0,0.2)', maxWidth: 320,
    }}>{msg}</div>
  );
}

// ── Styles constants ──────────────────────────────────────────────────────────

const S = {
  label: { display: 'block', fontSize: 12, fontWeight: 600, color: '#64748B', marginBottom: 4, textTransform: 'uppercase' as const, letterSpacing: '0.5px' },
  input: { width: '100%', padding: '8px 10px', border: '1px solid #CBD5E1', borderRadius: 6, fontSize: 14, color: '#1E293B', fontFamily: 'Inter, sans-serif', background: '#fff', boxSizing: 'border-box' as const },
  inputFocus: { outline: 'none', borderColor: '#1A2B4A' },
  select: { width: '100%', padding: '8px 10px', border: '1px solid #CBD5E1', borderRadius: 6, fontSize: 14, color: '#1E293B', fontFamily: 'Inter, sans-serif', background: '#fff', boxSizing: 'border-box' as const },
  section: { background: '#F8FAFC', border: '1px solid #E2E8F0', borderRadius: 8, padding: '16px', marginBottom: 16 },
  sectionTitle: { fontSize: 13, fontWeight: 700, color: '#1A2B4A', textTransform: 'uppercase' as const, letterSpacing: '0.5px', marginBottom: 12, paddingBottom: 8, borderBottom: '1px solid #E2E8F0' },
  btn: { padding: '10px 20px', borderRadius: 6, fontSize: 14, fontWeight: 600, cursor: 'pointer', border: 'none', fontFamily: 'Inter, sans-serif', transition: 'all 0.2s' },
  btnPrimary: { background: '#1A2B4A', color: '#fff' },
  btnSecondary: { background: '#F1F5F9', color: '#1E293B', border: '1px solid #CBD5E1' },
  btnDanger: { background: '#FEE2E2', color: '#DC2626', border: '1px solid #FECACA' },
  btnGreen: { background: '#D1FAE5', color: '#065F46', border: '1px solid #A7F3D0' },
  chip: (active: boolean) => ({
    width: 36, height: 36, display: 'flex', alignItems: 'center', justifyContent: 'center',
    borderRadius: 6, border: `2px solid ${active ? '#1A2B4A' : '#CBD5E1'}`,
    background: active ? '#1A2B4A' : '#fff', color: active ? '#fff' : '#64748B',
    fontSize: 13, fontWeight: 700, cursor: 'pointer', transition: 'all 0.15s', flexShrink: 0,
  }),
  badge: (s: string) => ({
    display: 'inline-flex', alignItems: 'center', padding: '2px 8px', borderRadius: 12, fontSize: 11, fontWeight: 600,
    background: s === 'saisie' ? '#DBEAFE' : s === 'facturee' ? '#D1FAE5' : s === 'emise' ? '#FEF3C7' : s === 'payee' ? '#D1FAE5' : '#F1F5F9',
    color: s === 'saisie' ? '#1E40AF' : s === 'facturee' ? '#065F46' : s === 'emise' ? '#92400E' : s === 'payee' ? '#065F46' : '#64748B',
  }),
  tableHeader: { background: '#1A2B4A', color: '#fff', padding: '10px 12px', fontSize: 12, fontWeight: 700, textAlign: 'left' as const, textTransform: 'uppercase' as const, letterSpacing: '0.5px', whiteSpace: 'nowrap' as const },
  tableCell: { padding: '10px 12px', fontSize: 13, color: '#1E293B', borderBottom: '1px solid #F1F5F9', verticalAlign: 'middle' as const },
};

// ── Module A: Saisie de fiche ─────────────────────────────────────────────────

function blankFiche() {
  return {
    assistant: '', numero_vol: '', type_vol: 'irregulier' as 'regulier' | 'irregulier',
    date_saisie: today(), site: 'LIBREVILLE', compagnie_assistee: COMPAGNIES[0],
    compagnie_autre: '', immatricule_aeronef: '',
    vol_national: false, vol_regional: false, vol_international: false,
    banques_depart: [] as number[], heure_ouverture_depart: '', date_ouverture_depart: today(),
    heure_cloture_depart: '', date_cloture_depart: today(), pax_prevu_depart: 0,
    banques_arrivee: [] as number[], heure_ouverture_arrivee: '', date_ouverture_arrivee: today(),
    heure_cloture_arrivee: '', date_cloture_arrivee: today(),
    pax_arrives: 0, pax_departs: 0, pax_transit: 0,
  };
}

function SaisieForm({ onSaved, editFiche, onToast }: { onSaved: () => void; editFiche: Fiche | null; onToast?: (msg: string) => void }) {
  const [f, setF] = useState(blankFiche());
  const [autreCompagnie, setAutreCompagnie] = useState(false);
  const [saving, setSaving] = useState(false);
  const [toast, setToast] = useState('');
  const showToast = (msg: string) => { setToast(msg); onToast?.(msg); };

  useEffect(() => {
    if (editFiche) {
      const isAutre = !COMPAGNIES.includes(editFiche.compagnie_assistee);
      setAutreCompagnie(isAutre);
      setF({
        assistant: editFiche.assistant, numero_vol: editFiche.numero_vol,
        type_vol: editFiche.type_vol, date_saisie: editFiche.date_saisie,
        site: editFiche.site,
        compagnie_assistee: isAutre ? 'autre' : editFiche.compagnie_assistee,
        compagnie_autre: isAutre ? editFiche.compagnie_assistee : '',
        immatricule_aeronef: editFiche.immatricule_aeronef,
        vol_national: editFiche.vol_national, vol_regional: editFiche.vol_regional,
        vol_international: editFiche.vol_international,
        banques_depart: editFiche.banques_depart, heure_ouverture_depart: editFiche.heure_ouverture_depart,
        date_ouverture_depart: editFiche.date_ouverture_depart, heure_cloture_depart: editFiche.heure_cloture_depart,
        date_cloture_depart: editFiche.date_cloture_depart, pax_prevu_depart: editFiche.pax_prevu_depart,
        banques_arrivee: editFiche.banques_arrivee, heure_ouverture_arrivee: editFiche.heure_ouverture_arrivee,
        date_ouverture_arrivee: editFiche.date_ouverture_arrivee, heure_cloture_arrivee: editFiche.heure_cloture_arrivee,
        date_cloture_arrivee: editFiche.date_cloture_arrivee,
        pax_arrives: editFiche.pax_arrives, pax_departs: editFiche.pax_departs, pax_transit: editFiche.pax_transit,
      });
    }
  }, [editFiche]);

  const set = (key: string, val: unknown) => setF(prev => ({ ...prev, [key]: val }));

  const dureeMin = calcDurationMinutes(f.heure_ouverture_arrivee, f.date_ouverture_arrivee, f.heure_cloture_arrivee, f.date_cloture_arrivee);
  const dureeH = roundToHalf(dureeMin);
  const totalPax = (f.pax_arrives || 0) + (f.pax_departs || 0) + (f.pax_transit || 0);

  const toggleBanque = (side: 'depart' | 'arrivee', n: number) => {
    const key = side === 'depart' ? 'banques_depart' : 'banques_arrivee';
    const arr = f[key] as number[];
    setF(prev => ({ ...prev, [key]: arr.includes(n) ? arr.filter(x => x !== n) : [...arr, n].sort((a, b) => a - b) }));
  };

  const compagnieValue = autreCompagnie ? f.compagnie_autre : f.compagnie_assistee;

  const handleSubmit = async () => {
    if (!f.numero_vol.trim()) { showToast('N° de vol obligatoire'); return; }
    if (!f.site) { showToast('Site obligatoire'); return; }
    if (!compagnieValue) { showToast('Compagnie obligatoire'); return; }
    setSaving(true);
    try {
      const payload = {
        ...f,
        compagnie_assistee: compagnieValue,
        immatricule_aeronef: f.immatricule_aeronef.toUpperCase(),
        nombre_banques_depart: f.banques_depart.length,
        nombre_banques_arrivee: f.banques_arrivee.length,
        duree_comptoirs_minutes: dureeMin,
        duree_heures_decimal: dureeH,
      };
      let result;
      if (editFiche) {
        result = await apiFetch('PUT', `/fiches-bandes/${editFiche.id}`, payload);
        if (!result.success) throw new Error(result.error || 'Erreur serveur');
        showToast('Fiche modifiée avec succès');
      } else {
        result = await apiFetch('POST', '/fiches-bandes', payload);
        if (!result.success) throw new Error(result.error || 'Erreur serveur');
        showToast(`Fiche ${result.numero_fiche} enregistrée avec succès !`);
        setF(blankFiche());
        setAutreCompagnie(false);
      }
      setTimeout(() => onSaved(), 1200);
    } catch (err: unknown) {
      const msg = err instanceof Error ? err.message : 'Connexion refusée — vérifiez que le serveur est lancé';
      showToast('Erreur : ' + msg);
      console.error('[BandesModule] Erreur sauvegarde fiche:', err);
    }
    setSaving(false);
  };

  const BanqueGrid = ({ side }: { side: 'depart' | 'arrivee' }) => {
    const arr = side === 'depart' ? f.banques_depart : f.banques_arrivee;
    return (
      <div style={{ display: 'flex', flexWrap: 'wrap', gap: 6, margin: '4px 0' }}>
        {[1,2,3,4,5,6,7,8,9,10].map(n => (
          <div key={n} style={S.chip(arr.includes(n))} onClick={() => toggleBanque(side, n)}>{n}</div>
        ))}
      </div>
    );
  };

  return (
    <div style={{ padding: '24px', maxWidth: 1000, margin: '0 auto' }}>
      {toast && <Toast msg={toast} onClose={() => setToast('')} />}

      {/* En-tête */}
      <div style={{ ...S.section, background: '#fff', border: '2px solid #1A2B4A' }}>
        <div style={{ textAlign: 'center', marginBottom: 16 }}>
          <div style={{ fontSize: 13, color: '#64748B' }}>ASECNA — Délégation aux Activités Aéronautiques Nationales du Gabon</div>
          <div style={{ fontSize: 16, fontWeight: 700, color: '#1A2B4A', marginTop: 4 }}>FICHE DE GESTION DES COMPTOIRS D'ENREGISTREMENT</div>
          <div style={{ fontSize: 22, fontWeight: 800, color: '#DC2626', marginTop: 6 }}>
            N° {editFiche ? editFiche.numero_fiche : '(auto)'}
          </div>
        </div>
        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr 1fr', gap: 12 }}>
          <div>
            <label style={S.label}>Assistant</label>
            <input style={S.input} value={f.assistant} onChange={e => set('assistant', e.target.value)} placeholder="Nom de l'assistant" />
          </div>
          <div>
            <label style={S.label}>N° Vol</label>
            <input style={S.input} value={f.numero_vol} onChange={e => set('numero_vol', e.target.value.toUpperCase())} placeholder="ex: SPA911" />
          </div>
          <div>
            <label style={S.label}>Type de vol</label>
            <div style={{ display: 'flex', gap: 16, marginTop: 8 }}>
              {(['regulier', 'irregulier'] as const).map(t => (
                <label key={t} style={{ display: 'flex', alignItems: 'center', gap: 6, cursor: 'pointer', fontSize: 13 }}>
                  <input type="radio" name="type_vol" value={t} checked={f.type_vol === t} onChange={() => set('type_vol', t)} />
                  {t === 'regulier' ? 'Régulier' : 'Irrégulier'}
                </label>
              ))}
            </div>
          </div>
          <div>
            <label style={S.label}>Date</label>
            <input style={S.input} type="date" value={f.date_saisie} onChange={e => set('date_saisie', e.target.value)} />
          </div>
          <div>
            <label style={S.label}>Site</label>
            <select style={S.select} value={f.site} onChange={e => set('site', e.target.value)}>
              {SITES.map(s => <option key={s}>{s}</option>)}
              <option value="autre">Autre</option>
            </select>
          </div>
          <div>
            <label style={S.label}>Compagnie assistée</label>
            <select style={S.select} value={autreCompagnie ? 'autre' : f.compagnie_assistee}
              onChange={e => { if (e.target.value === 'autre') { setAutreCompagnie(true); set('compagnie_assistee', 'autre'); } else { setAutreCompagnie(false); set('compagnie_assistee', e.target.value); } }}>
              {COMPAGNIES.map(c => <option key={c}>{c}</option>)}
              <option value="autre">Autre...</option>
            </select>
            {autreCompagnie && <input style={{ ...S.input, marginTop: 6 }} value={f.compagnie_autre} onChange={e => set('compagnie_autre', e.target.value.toUpperCase())} placeholder="Nom de la compagnie" />}
          </div>
          <div>
            <label style={S.label}>Immatricule aéronef</label>
            <input style={S.input} value={f.immatricule_aeronef} onChange={e => set('immatricule_aeronef', e.target.value.toUpperCase())} placeholder="ex: TR-ABX" />
          </div>
          <div style={{ gridColumn: '2 / 4' }}>
            <label style={S.label}>Type de vol</label>
            <div style={{ display: 'flex', gap: 20, marginTop: 6 }}>
              {[['vol_national','National'],['vol_regional','Régional'],['vol_international','International']].map(([k,l]) => (
                <label key={k} style={{ display: 'flex', alignItems: 'center', gap: 6, cursor: 'pointer', fontSize: 13 }}>
                  <input type="checkbox" checked={!!(f as Record<string,unknown>)[k]} onChange={e => set(k, e.target.checked)} />
                  {l}
                </label>
              ))}
            </div>
          </div>
        </div>
      </div>

      {/* Sections Départ / Arrivée */}
      <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 16 }}>
        {/* DÉPART */}
        <div style={S.section}>
          <div style={S.sectionTitle}>Départ</div>
          <div style={{ marginBottom: 10 }}>
            <label style={S.label}>Banques utilisées (cliquer pour sélectionner)</label>
            <BanqueGrid side="depart" />
            <div style={{ fontSize: 12, color: '#64748B', marginTop: 4 }}>
              {f.banques_depart.length > 0 ? `Sélectionné : ${f.banques_depart.join(', ')} (${f.banques_depart.length} banque${f.banques_depart.length > 1 ? 's' : ''})` : 'Aucune banque sélectionnée'}
            </div>
          </div>
          <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 8 }}>
            <div>
              <label style={S.label}>Heure d'ouverture</label>
              <input style={S.input} type="time" value={f.heure_ouverture_depart} onChange={e => set('heure_ouverture_depart', e.target.value)} />
            </div>
            <div>
              <label style={S.label}>Date d'ouverture</label>
              <input style={S.input} type="date" value={f.date_ouverture_depart} onChange={e => set('date_ouverture_depart', e.target.value)} />
            </div>
            <div>
              <label style={S.label}>Heure de clôture</label>
              <input style={S.input} type="time" value={f.heure_cloture_depart} onChange={e => set('heure_cloture_depart', e.target.value)} />
            </div>
            <div>
              <label style={S.label}>Date de clôture</label>
              <input style={S.input} type="date" value={f.date_cloture_depart} onChange={e => set('date_cloture_depart', e.target.value)} />
            </div>
          </div>
          <div style={{ marginTop: 8 }}>
            <label style={S.label}>Pax prévus départ</label>
            <input style={S.input} type="number" min="0" value={f.pax_prevu_depart || ''} onChange={e => set('pax_prevu_depart', parseInt(e.target.value) || 0)} />
          </div>
        </div>

        {/* ARRIVÉE */}
        <div style={S.section}>
          <div style={S.sectionTitle}>Arrivée</div>
          <div style={{ marginBottom: 10 }}>
            <label style={S.label}>Banques utilisées (cliquer pour sélectionner)</label>
            <BanqueGrid side="arrivee" />
            <div style={{ fontSize: 12, color: '#64748B', marginTop: 4 }}>
              {f.banques_arrivee.length > 0 ? `Sélectionné : ${f.banques_arrivee.join(', ')} (${f.banques_arrivee.length} banque${f.banques_arrivee.length > 1 ? 's' : ''})` : 'Aucune banque sélectionnée'}
            </div>
          </div>
          <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 8 }}>
            <div>
              <label style={S.label}>Heure d'ouverture</label>
              <input style={S.input} type="time" value={f.heure_ouverture_arrivee} onChange={e => set('heure_ouverture_arrivee', e.target.value)} />
            </div>
            <div>
              <label style={S.label}>Date d'ouverture</label>
              <input style={S.input} type="date" value={f.date_ouverture_arrivee} onChange={e => set('date_ouverture_arrivee', e.target.value)} />
            </div>
            <div>
              <label style={S.label}>Heure de clôture</label>
              <input style={S.input} type="time" value={f.heure_cloture_arrivee} onChange={e => set('heure_cloture_arrivee', e.target.value)} />
            </div>
            <div>
              <label style={S.label}>Date de clôture</label>
              <input style={S.input} type="date" value={f.date_cloture_arrivee} onChange={e => set('date_cloture_arrivee', e.target.value)} />
            </div>
          </div>

          {/* Durée auto */}
          <div style={{ marginTop: 8, background: '#EFF6FF', border: '1px solid #BFDBFE', borderRadius: 6, padding: '10px 14px' }}>
            <div style={{ fontSize: 12, fontWeight: 700, color: '#1E40AF', marginBottom: 2 }}>DURÉE COMPTOIRS (auto-calculée)</div>
            {dureeMin > 0 ? (
              <div style={{ fontSize: 15, fontWeight: 700, color: '#1A2B4A' }}>
                {minsToHHMM(dureeMin)} = <span style={{ color: '#1E40AF' }}>{dureeH}h facturées</span>
              </div>
            ) : <div style={{ color: '#94A3B8', fontSize: 13 }}>Saisir les heures pour calculer</div>}
          </div>

          <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr 1fr', gap: 8, marginTop: 8 }}>
            <div>
              <label style={S.label}>PAX Arrivée</label>
              <input style={S.input} type="number" min="0" value={f.pax_arrives || ''} onChange={e => set('pax_arrives', parseInt(e.target.value) || 0)} />
            </div>
            <div>
              <label style={S.label}>PAX Départ</label>
              <input style={S.input} type="number" min="0" value={f.pax_departs || ''} onChange={e => set('pax_departs', parseInt(e.target.value) || 0)} />
            </div>
            <div>
              <label style={S.label}>Transit</label>
              <input style={S.input} type="number" min="0" value={f.pax_transit || ''} onChange={e => set('pax_transit', parseInt(e.target.value) || 0)} />
            </div>
          </div>
        </div>
      </div>

      {/* Résumé facturation */}
      {dureeH > 0 && (
        <div style={{ background: '#F0FDF4', border: '1px solid #BBF7D0', borderRadius: 8, padding: '14px 20px', marginBottom: 16, display: 'flex', gap: 32, alignItems: 'center' }}>
          <div>
            <div style={{ fontSize: 12, color: '#64748B', fontWeight: 600 }}>DURÉE FACTURÉE</div>
            <div style={{ fontSize: 18, fontWeight: 800, color: '#065F46' }}>{dureeH}h</div>
          </div>
          <div>
            <div style={{ fontSize: 12, color: '#64748B', fontWeight: 600 }}>MONTANT ESTIMÉ</div>
            <div style={{ fontSize: 18, fontWeight: 800, color: '#065F46' }}>{fmt(dureeH * 10000)} FCFA</div>
          </div>
          <div>
            <div style={{ fontSize: 12, color: '#64748B', fontWeight: 600 }}>TOTAL PAX</div>
            <div style={{ fontSize: 18, fontWeight: 800, color: '#065F46' }}>{totalPax}</div>
          </div>
        </div>
      )}

      <div style={{ display: 'flex', gap: 12 }}>
        <button style={{ ...S.btn, ...S.btnPrimary }} onClick={handleSubmit} disabled={saving}>
          {saving ? 'Enregistrement...' : editFiche ? 'Mettre à jour la fiche' : 'Enregistrer la fiche'}
        </button>
        <button style={{ ...S.btn, ...S.btnSecondary }} onClick={() => { setF(blankFiche()); setAutreCompagnie(false); }}>
          Réinitialiser
        </button>
      </div>
    </div>
  );
}

// ── Module B: Liste des fiches ────────────────────────────────────────────────

function FichesList({
  onEdit, onFacturer, onRefresh,
}: { onEdit: (f: Fiche) => void; onFacturer: (ids: string[]) => void; onRefresh: number }) {
  const [fiches, setFiches] = useState<Fiche[]>([]);
  const [loading, setLoading] = useState(false);
  const [selected, setSelected] = useState<Set<string>>(new Set());
  const [filters, setFilters] = useState({ compagnie: '', site: '', statut: '', search: '' });
  const [toast, setToast] = useState('');

  const load = useCallback(async () => {
    setLoading(true);
    const params = new URLSearchParams();
    if (filters.compagnie) params.append('compagnie', filters.compagnie);
    if (filters.site) params.append('site', filters.site);
    if (filters.statut) params.append('statut', filters.statut);
    const data = await apiFetch('GET', `/fiches-bandes?${params}`);
    if (data.success) setFiches(data.data);
    setLoading(false);
  }, [filters, onRefresh]);

  useEffect(() => { load(); }, [load]);

  const displayed = fiches.filter(f => {
    if (!filters.search) return true;
    const s = filters.search.toLowerCase();
    return f.numero_fiche.includes(s) || f.numero_vol.toLowerCase().includes(s) ||
      f.compagnie_assistee.toLowerCase().includes(s) || f.assistant.toLowerCase().includes(s);
  });

  const toggleSelect = (id: string) => {
    setSelected(prev => {
      const next = new Set(prev);
      next.has(id) ? next.delete(id) : next.add(id);
      return next;
    });
  };

  const toggleAll = () => {
    if (selected.size === displayed.length) setSelected(new Set());
    else setSelected(new Set(displayed.map(f => f.id)));
  };

  const handleDelete = async (id: string) => {
    if (!confirm('Supprimer cette fiche ?')) return;
    await apiFetch('DELETE', `/fiches-bandes/${id}`);
    setToast('Fiche supprimée');
    load();
  };

  const handleFacturerSelection = () => {
    if (selected.size === 0) { setToast('Sélectionnez au moins une fiche'); return; }
    const selectedFiches = fiches.filter(f => selected.has(f.id));
    const companies = new Set(selectedFiches.map(f => f.compagnie_assistee));
    if (companies.size > 1) { setToast('Toutes les fiches doivent appartenir à la même compagnie'); return; }
    const alreadyFactured = selectedFiches.filter(f => f.statut === 'facturee');
    if (alreadyFactured.length > 0) { setToast(`${alreadyFactured.length} fiche(s) déjà facturée(s)`); return; }
    onFacturer(Array.from(selected));
  };

  return (
    <div style={{ padding: 24 }}>
      {toast && <Toast msg={toast} onClose={() => setToast('')} />}

      {/* Filtres */}
      <div style={{ display: 'flex', gap: 12, marginBottom: 16, flexWrap: 'wrap' }}>
        <input style={{ ...S.input, width: 200 }} placeholder="Rechercher..." value={filters.search}
          onChange={e => setFilters(p => ({ ...p, search: e.target.value }))} />
        <select style={{ ...S.select, width: 180 }} value={filters.compagnie}
          onChange={e => setFilters(p => ({ ...p, compagnie: e.target.value }))}>
          <option value="">Toutes compagnies</option>
          {COMPAGNIES.map(c => <option key={c}>{c}</option>)}
        </select>
        <select style={{ ...S.select, width: 160 }} value={filters.site}
          onChange={e => setFilters(p => ({ ...p, site: e.target.value }))}>
          <option value="">Tous sites</option>
          {SITES.map(s => <option key={s}>{s}</option>)}
        </select>
        <select style={{ ...S.select, width: 150 }} value={filters.statut}
          onChange={e => setFilters(p => ({ ...p, statut: e.target.value }))}>
          <option value="">Tous statuts</option>
          <option value="saisie">Saisie</option>
          <option value="facturee">Facturée</option>
        </select>
      </div>

      {selected.size > 0 && (
        <div style={{ background: '#EFF6FF', border: '1px solid #BFDBFE', borderRadius: 8, padding: '10px 16px', marginBottom: 12, display: 'flex', alignItems: 'center', gap: 12 }}>
          <span style={{ fontSize: 13, color: '#1E40AF', fontWeight: 600 }}>{selected.size} fiche(s) sélectionnée(s)</span>
          <button style={{ ...S.btn, ...S.btnPrimary, padding: '6px 16px', fontSize: 13 }} onClick={handleFacturerSelection}>
            Créer une facture depuis la sélection
          </button>
          <button style={{ ...S.btn, ...S.btnSecondary, padding: '6px 12px', fontSize: 13 }} onClick={() => setSelected(new Set())}>
            Annuler
          </button>
        </div>
      )}

      {loading ? (
        <div style={{ textAlign: 'center', padding: 40, color: '#64748B' }}>Chargement...</div>
      ) : (
        <div style={{ overflowX: 'auto' }}>
          <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 13 }}>
            <thead>
              <tr>
                <th style={{ ...S.tableHeader, width: 40 }}>
                  <input type="checkbox" checked={selected.size === displayed.length && displayed.length > 0}
                    onChange={toggleAll} />
                </th>
                <th style={S.tableHeader}>N° Fiche</th>
                <th style={S.tableHeader}>Date</th>
                <th style={S.tableHeader}>Assistant</th>
                <th style={S.tableHeader}>Compagnie</th>
                <th style={S.tableHeader}>Site</th>
                <th style={S.tableHeader}>Vol</th>
                <th style={S.tableHeader}>Durée (h)</th>
                <th style={S.tableHeader}>PAX</th>
                <th style={S.tableHeader}>Statut</th>
                <th style={S.tableHeader}>Actions</th>
              </tr>
            </thead>
            <tbody>
              {displayed.length === 0 ? (
                <tr><td colSpan={11} style={{ ...S.tableCell, textAlign: 'center', color: '#94A3B8', padding: 32 }}>
                  Aucune fiche trouvée
                </td></tr>
              ) : displayed.map((fiche, i) => (
                <tr key={fiche.id} style={{ background: i % 2 === 0 ? '#fff' : '#F8FAFC' }}>
                  <td style={S.tableCell}>
                    <input type="checkbox" checked={selected.has(fiche.id)} onChange={() => toggleSelect(fiche.id)} />
                  </td>
                  <td style={{ ...S.tableCell, fontWeight: 700, color: '#1A2B4A' }}>{fiche.numero_fiche}</td>
                  <td style={S.tableCell}>{fiche.date_saisie}</td>
                  <td style={S.tableCell}>{fiche.assistant || '—'}</td>
                  <td style={S.tableCell}>{fiche.compagnie_assistee}</td>
                  <td style={S.tableCell}>{fiche.site}</td>
                  <td style={S.tableCell}>{fiche.numero_vol}</td>
                  <td style={{ ...S.tableCell, fontWeight: 600 }}>{fiche.duree_heures_decimal}h</td>
                  <td style={S.tableCell}>{(fiche.pax_arrives || 0) + (fiche.pax_departs || 0) + (fiche.pax_transit || 0)}</td>
                  <td style={S.tableCell}><span style={S.badge(fiche.statut)}>{fiche.statut}</span></td>
                  <td style={S.tableCell}>
                    <div style={{ display: 'flex', gap: 4 }}>
                      <button style={{ ...S.btn, ...S.btnSecondary, padding: '4px 10px', fontSize: 12 }} onClick={() => onEdit(fiche)}>Modifier</button>
                      {fiche.statut === 'saisie' && (
                        <button style={{ ...S.btn, ...S.btnGreen, padding: '4px 10px', fontSize: 12 }} onClick={() => onFacturer([fiche.id])}>Facturer</button>
                      )}
                      <button style={{ ...S.btn, ...S.btnDanger, padding: '4px 10px', fontSize: 12 }} onClick={() => handleDelete(fiche.id)}>✕</button>
                    </div>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}

// ── Excel generation — basée sur le template BK ───────────────────────────────

async function generateFactureExcel(facture: Facture) {
  // 1. Charger le template depuis public/
  const res = await fetch("/Facturation bandes d'enregistrements de 2026-V1.xlsx");
  const arrayBuffer = await res.arrayBuffer();
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(arrayBuffer);

  // 2. Récupérer la feuille BK
  const ws = wb.worksheets.find(s => s.name.toLowerCase() === 'bk');
  if (!ws) throw new Error('Feuille BK introuvable dans le template');

  // 3. Formater les dates pour la période (DD/MM sans année, DD/MM/YYYY pour fin)
  const fmtShort = (iso: string) => {
    if (!iso) return '';
    const [, m, d] = iso.split('-');
    return `${d}/${m}`;
  };
  const fmtFull = (iso: string) => {
    if (!iso) return '';
    const [y, m, d] = iso.split('-');
    return `${d}/${m}/${y}`;
  };
  const periodeStr = (facture.periode_debut && facture.periode_fin)
    ? `Du ${fmtShort(facture.periode_debut)} au ${fmtFull(facture.periode_fin)}`
    : '';

  // 4. Remplir le bloc 1 (cols A–H, lignes 1–39) — garder tous les styles du template
  ws.getCell('B4').value = `Facture N°${facture.numero_facture}`;

  // Compagnie (F7:H7 mergé)
  ws.getCell('F7').value = facture.compagnie;
  // "Libreville, le {date}" en A8
  ws.getCell('A8').value = `Libreville, le ${new Date(facture.date_facture + 'T00:00:00').toLocaleDateString('fr-FR')}`;
  // Adresse compagnie (F8:H8 mergé)
  ws.getCell('F8').value = facture.adresse_compagnie || '';
  // Ville compagnie (F9:H9 mergé)
  ws.getCell('F9').value = facture.ville_compagnie || facture.site;

  // Site et Série
  ws.getCell('C11').value = facture.site;
  ws.getCell('B12').value = `Série N°:${facture.serie_bandes || ''}`;

  // Période (A16 mergé A16:A22)
  ws.getCell('A16').value = periodeStr;

  // Données tableau ligne 19
  ws.getCell('B19').value = facture.nombre_heures;
  ws.getCell('C19').value = facture.tarif_horaire;
  ws.getCell('D19').value = facture.total_heures;
  // Annonces : afficher 0 si pas d'annonces (le template gère le "-" via la valeur 0)
  ws.getCell('E19').value = facture.nombre_annonces;
  ws.getCell('F19').value = facture.nombre_annonces === 0 ? 0 : facture.tarif_annonce;
  ws.getCell('G19').value = facture.total_annonces;
  ws.getCell('H19').value = facture.montant_ht;

  // Totaux ligne 24
  ws.getCell('A24').value = facture.total_pax;
  ws.getCell('B24').value = facture.montant_ht;
  ws.getCell('D24').value = facture.taxes;
  ws.getCell('F24').value = facture.acompte;
  ws.getCell('G24').value = facture.montant_ht;
  ws.getCell('H24').value = facture.solde;

  // Montant en lettres (A27 mergé A27:G27)
  ws.getCell('A27').value = facture.montant_en_lettres;

  // 5. Effacer le bloc 2 (cols I–P, lignes 1–39) — données de l'exemple
  const bloc2DataCells = [
    'J4',                         // N° facture 2
    'N7',                         // Compagnie 2
    'I8', 'N8',                   // "Libreville" + adresse 2
    'N9',                         // Ville 2
    'K11', 'J12',                 // Site + Série 2
    'I16',                        // Période 2
    'J19', 'K19', 'L19', 'M19', 'N19', 'O19', 'P19',  // Données tableau 2
    'I23', 'J23', 'L23', 'N23', 'O23', 'P23',          // Headers totaux 2
    'I24', 'J24', 'L24', 'N24', 'O24', 'P24',          // Totaux 2
    'I26', 'I27',                 // Certifié + montant lettres 2
  ];
  for (const addr of bloc2DataCells) {
    try { ws.getCell(addr).value = null; } catch { /* cellule protégée */ }
  }

  // 6. Effacer les blocs 3 et 4 (lignes 41–79)
  for (let r = 41; r <= ws.rowCount; r++) {
    ws.getRow(r).eachCell({ includeEmpty: false }, cell => {
      try { cell.value = null; } catch { /* skip */ }
    });
  }

  // 7. Télécharger
  const buf = await wb.xlsx.writeBuffer();
  const blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  saveAs(blob, `Facture-Bandes-${facture.numero_facture}-${facture.compagnie}.xlsx`);
}

// ── Module C: Génération de facture ──────────────────────────────────────────

function FacturationForm({ ficheIds, fiches, onSaved }: { ficheIds: string[]; fiches: Fiche[]; onSaved: () => void }) {
  const selectedFiches = fiches.filter(f => ficheIds.includes(f.id));
  const totalH = selectedFiches.reduce((s, f) => s + (f.duree_heures_decimal || 0), 0);
  const totalPax = selectedFiches.reduce((s, f) => s + (f.pax_arrives || 0) + (f.pax_departs || 0) + (f.pax_transit || 0), 0);
  const defaultSite = selectedFiches[0]?.site || '';
  const defaultComp = selectedFiches[0]?.compagnie_assistee || '';

  const [form, setForm] = useState({
    date_facture: today(),
    compagnie: defaultComp,
    adresse_compagnie: '',
    ville_compagnie: '',
    site: defaultSite,
    serie_bandes: '',
    periode_debut: selectedFiches.length > 0 ? selectedFiches[selectedFiches.length - 1].date_saisie : today(),
    periode_fin: selectedFiches.length > 0 ? selectedFiches[0].date_saisie : today(),
    nombre_heures: Math.round(totalH * 2) / 2,
    tarif_horaire: 10000,
    nombre_annonces: 0,
    tarif_annonce: 3500,
    taxes: 0,
    acompte: 0,
    total_pax: totalPax,
  });
  const [saving, setSaving] = useState(false);
  const [toast, setToast] = useState('');
  const [generatingExcel, setGeneratingExcel] = useState(false);

  const set = (k: string, v: unknown) => setForm(p => ({ ...p, [k]: v }));

  const total_heures = form.nombre_heures * form.tarif_horaire;
  const total_annonces = form.nombre_annonces * form.tarif_annonce;
  const montant_ht = total_heures + total_annonces;
  const solde = montant_ht + form.taxes - form.acompte;
  const montant_en_lettres = numberToWords(solde);

  const buildPayload = (statut: string) => ({
    ...form,
    fiches_ids: ficheIds,
    total_heures,
    total_annonces,
    montant_ht,
    solde,
    montant_en_lettres,
    statut,
  });

  const handleSave = async (statut: string) => {
    if (!form.compagnie) { setToast('Compagnie obligatoire'); return; }
    setSaving(true);
    try {
      const data = await apiFetch('POST', '/factures-bandes', buildPayload(statut));
      if (data.success) {
        setToast(`Facture N°${data.numero_facture} enregistrée`);
        onSaved();
      } else setToast('Erreur: ' + data.error);
    } catch { setToast('Erreur de sauvegarde'); }
    setSaving(false);
  };

  const handleExcel = async () => {
    if (!form.compagnie) { setToast('Compagnie obligatoire'); return; }
    setGeneratingExcel(true);
    try {
      const tempFacture: Facture = {
        id: 'preview', numero_facture: '????',
        ...form, fiches_ids: ficheIds,
        total_heures, total_annonces, montant_ht, solde, montant_en_lettres,
        statut: 'brouillon',
      };
      await generateFactureExcel(tempFacture);
      setToast('Fichier Excel généré');
    } catch (e) { setToast('Erreur génération Excel'); }
    setGeneratingExcel(false);
  };

  return (
    <div style={{ padding: 24, maxWidth: 800, margin: '0 auto' }}>
      {toast && <Toast msg={toast} onClose={() => setToast('')} />}

      <div style={S.section}>
        <div style={S.sectionTitle}>Informations de la facture</div>
        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 12 }}>
          <div>
            <label style={S.label}>Date de facture</label>
            <input style={S.input} type="date" value={form.date_facture} onChange={e => set('date_facture', e.target.value)} />
          </div>
          <div>
            <label style={S.label}>Compagnie</label>
            <select style={S.select} value={form.compagnie} onChange={e => set('compagnie', e.target.value)}>
              {COMPAGNIES.map(c => <option key={c}>{c}</option>)}
            </select>
          </div>
          <div>
            <label style={S.label}>Adresse compagnie (BP + Tél)</label>
            <input style={S.input} value={form.adresse_compagnie} onChange={e => set('adresse_compagnie', e.target.value)} placeholder="ex: BP:13025 Tél:011 44 40 15" />
          </div>
          <div>
            <label style={S.label}>Ville compagnie</label>
            <input style={S.input} value={form.ville_compagnie} onChange={e => set('ville_compagnie', e.target.value)} placeholder="ex: Oyem" />
          </div>
          <div>
            <label style={S.label}>Site</label>
            <select style={S.select} value={form.site} onChange={e => set('site', e.target.value)}>
              {SITES.map(s => <option key={s}>{s}</option>)}
            </select>
          </div>
          <div>
            <label style={S.label}>Série N° des bandes</label>
            <input style={S.input} value={form.serie_bandes} onChange={e => set('serie_bandes', e.target.value)} placeholder="ex: 0001451-0001491" />
          </div>
          <div>
            <label style={S.label}>Période — Du</label>
            <input style={S.input} type="date" value={form.periode_debut} onChange={e => set('periode_debut', e.target.value)} />
          </div>
          <div>
            <label style={S.label}>Période — Au</label>
            <input style={S.input} type="date" value={form.periode_fin} onChange={e => set('periode_fin', e.target.value)} />
          </div>
        </div>
        {ficheIds.length > 0 && (
          <div style={{ marginTop: 12, padding: '10px 14px', background: '#EFF6FF', borderRadius: 6, fontSize: 13, color: '#1E40AF' }}>
            {ficheIds.length} fiche(s) incluse(s) • {selectedFiches.map(f => f.numero_fiche).join(', ')}
          </div>
        )}
      </div>

      <div style={S.section}>
        <div style={S.sectionTitle}>Calcul de la facture</div>
        <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 13 }}>
          <thead>
            <tr>
              <th style={{ ...S.tableHeader, textAlign: 'center' }} colSpan={4}>Usage des Comptoirs</th>
              <th style={{ ...S.tableHeader, textAlign: 'center' }} colSpan={4}>Annonces point "I"</th>
            </tr>
            <tr>
              {['Nbre H','CU (FCFA)','Total 1','','Nbre Ann','CU (FCFA)','Total 2','Montant HT'].map((h, i) => (
                <th key={i} style={{ ...S.tableHeader, background: '#E2E8F0', color: '#1E293B', fontSize: 12, padding: '8px 10px' }}>{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            <tr>
              <td style={S.tableCell}><input style={{ ...S.input, width: 80, textAlign: 'center' }} type="number" min="0" step="0.5" value={form.nombre_heures} onChange={e => set('nombre_heures', parseFloat(e.target.value) || 0)} /></td>
              <td style={S.tableCell}><input style={{ ...S.input, width: 90, textAlign: 'center' }} type="number" value={form.tarif_horaire} onChange={e => set('tarif_horaire', parseInt(e.target.value) || 0)} /></td>
              <td style={{ ...S.tableCell, fontWeight: 700 }}>{fmt(total_heures)}</td>
              <td style={S.tableCell}></td>
              <td style={S.tableCell}><input style={{ ...S.input, width: 80, textAlign: 'center' }} type="number" min="0" value={form.nombre_annonces} onChange={e => set('nombre_annonces', parseInt(e.target.value) || 0)} /></td>
              <td style={S.tableCell}><input style={{ ...S.input, width: 90, textAlign: 'center' }} type="number" value={form.tarif_annonce} onChange={e => set('tarif_annonce', parseInt(e.target.value) || 0)} /></td>
              <td style={{ ...S.tableCell, fontWeight: 700 }}>{fmt(total_annonces)}</td>
              <td style={{ ...S.tableCell, fontWeight: 700, fontSize: 15, color: '#1A2B4A' }}>{fmt(montant_ht)}</td>
            </tr>
          </tbody>
        </table>

        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr 1fr 1fr', gap: 12, marginTop: 16 }}>
          <div>
            <label style={S.label}>PAX Total</label>
            <input style={S.input} type="number" value={form.total_pax} onChange={e => set('total_pax', parseInt(e.target.value) || 0)} />
          </div>
          <div>
            <label style={S.label}>Taxes</label>
            <input style={S.input} type="number" value={form.taxes} onChange={e => set('taxes', parseInt(e.target.value) || 0)} />
          </div>
          <div>
            <label style={S.label}>Acompte</label>
            <input style={S.input} type="number" value={form.acompte} onChange={e => set('acompte', parseInt(e.target.value) || 0)} />
          </div>
          <div>
            <label style={S.label}>SOLDE À PAYER</label>
            <div style={{ padding: '8px 10px', background: '#1A2B4A', color: '#fff', borderRadius: 6, fontWeight: 800, fontSize: 16 }}>{fmt(solde)} FCFA</div>
          </div>
        </div>

        <div style={{ marginTop: 12, padding: '10px 14px', background: '#F8FAFC', border: '1px solid #E2E8F0', borderRadius: 6 }}>
          <div style={{ fontSize: 11, color: '#64748B', fontWeight: 600, marginBottom: 4 }}>MONTANT EN LETTRES</div>
          <div style={{ fontSize: 13, fontWeight: 600, color: '#1A2B4A' }}>{montant_en_lettres}</div>
        </div>
      </div>

      <div style={{ display: 'flex', gap: 12, flexWrap: 'wrap' }}>
        <button style={{ ...S.btn, ...S.btnPrimary }} onClick={() => handleSave('brouillon')} disabled={saving}>
          Enregistrer en brouillon
        </button>
        <button style={{ ...S.btn, ...S.btnGreen }} onClick={() => handleSave('emise')} disabled={saving}>
          Marquer comme émise
        </button>
        <button style={{ ...S.btn, ...S.btnSecondary }} onClick={handleExcel} disabled={generatingExcel}>
          {generatingExcel ? 'Génération...' : '📥 Générer Excel'}
        </button>
      </div>
    </div>
  );
}

// ── Module D: Bordereau ───────────────────────────────────────────────────────

async function generateBordereauExcel(factures: Facture[]) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Recap');
  ws.columns = [
    { width: 12 }, { width: 16 }, { width: 20 }, { width: 22 }, { width: 14 }, { width: 16 }, { width: 12 },
  ];
  const blue = { type: 'pattern' as const, pattern: 'solid' as const, fgColor: { argb: 'FF1A2B4A' } };
  const border = { top: { style: 'thin' as const }, bottom: { style: 'thin' as const }, left: { style: 'thin' as const }, right: { style: 'thin' as const } };

  ws.mergeCells('A1:G1');
  ws.getCell('A1').value = "BORDEREAU D'EMISSION DE BANQUES D'ENREGISTREMENT";
  ws.getCell('A1').font = { bold: true, size: 13, color: { argb: 'FFFFFFFF' } };
  ws.getCell('A1').fill = blue;
  ws.getCell('A1').alignment = { horizontal: 'center', vertical: 'middle' };
  ws.getRow(1).height = 28;

  const headers = ['N° FACTURE', 'DATE FACTURE', 'CLIENT', 'PÉRIODE', 'SITE', 'MONTANT (FCFA)', 'STATUT'];
  const hRow = ws.addRow(headers);
  hRow.height = 22;
  hRow.eachCell(cell => {
    cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2C3E50' } };
    cell.border = border;
    cell.alignment = { horizontal: 'center', vertical: 'middle' };
  });

  let total = 0;
  for (const f of factures) {
    const periode = f.periode_debut && f.periode_fin
      ? `${new Date(f.periode_debut).toLocaleDateString('fr-FR')} — ${new Date(f.periode_fin).toLocaleDateString('fr-FR')}`
      : '—';
    const row = ws.addRow([f.numero_facture, new Date(f.date_facture).toLocaleDateString('fr-FR'), f.compagnie, periode, f.site, f.montant_ht, f.statut]);
    row.getCell(6).numFmt = '#,##0';
    row.eachCell(cell => { cell.border = border; cell.alignment = { vertical: 'middle' }; });
    total += f.montant_ht;
  }

  const totalRow = ws.addRow(['TOTAL', '', '', '', '', total, '']);
  totalRow.getCell(1).font = { bold: true };
  totalRow.getCell(6).font = { bold: true };
  totalRow.getCell(6).numFmt = '#,##0';
  totalRow.eachCell(cell => { cell.border = border; cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFBDD7EE' } }; });

  const buf = await wb.xlsx.writeBuffer();
  saveAs(new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), `Bordereau-Bandes-${new Date().getFullYear()}.xlsx`);
}

function Bordereau() {
  const [factures, setFactures] = useState<Facture[]>([]);
  const [loading, setLoading] = useState(false);
  const [toast, setToast] = useState('');
  const [stats, setStats] = useState<{ fichesTotal: number; fichesSaisie: number; facturesMois: number; facturesAttente: number } | null>(null);

  const load = useCallback(async () => {
    setLoading(true);
    const [fData, sData] = await Promise.all([
      apiFetch('GET', '/factures-bandes'),
      apiFetch('GET', '/stats-bandes'),
    ]);
    if (fData.success) setFactures(fData.data);
    if (sData.success) setStats(sData.data);
    setLoading(false);
  }, []);

  useEffect(() => { load(); }, [load]);

  const total = factures.reduce((s, f) => s + (f.montant_ht || 0), 0);

  const handleDelete = async (id: string) => {
    if (!confirm('Supprimer cette facture et libérer les fiches associées ?')) return;
    await apiFetch('DELETE', `/factures-bandes/${id}`);
    setToast('Facture supprimée');
    load();
  };

  const handleStatusChange = async (f: Facture, statut: string) => {
    await apiFetch('PUT', `/factures-bandes/${f.id}`, { ...f, statut });
    load();
  };

  return (
    <div style={{ padding: 24 }}>
      {toast && <Toast msg={toast} onClose={() => setToast('')} />}

      {/* KPIs */}
      {stats && (
        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(4, 1fr)', gap: 12, marginBottom: 20 }}>
          {[
            ['Fiches du mois', stats.fichesTotal, '#EFF6FF', '#1E40AF'],
            ['Total facturé mois', fmt(stats.facturesMois) + ' FCFA', '#F0FDF4', '#065F46'],
            ['Fiches non facturées', stats.fichesSaisie, '#FEF3C7', '#92400E'],
            ['Factures en attente', stats.facturesAttente, '#FEE2E2', '#DC2626'],
          ].map(([label, val, bg, color]) => (
            <div key={label as string} style={{ background: bg as string, border: `1px solid ${color as string}33`, borderRadius: 8, padding: '16px 20px' }}>
              <div style={{ fontSize: 11, fontWeight: 700, color: color as string, textTransform: 'uppercase', letterSpacing: '0.5px', marginBottom: 8 }}>{label as string}</div>
              <div style={{ fontSize: 22, fontWeight: 800, color: color as string }}>{val as string | number}</div>
            </div>
          ))}
        </div>
      )}

      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 16 }}>
        <div style={{ fontWeight: 700, fontSize: 16, color: '#1A2B4A' }}>
          Bordereau — {factures.length} facture(s) — Total : {fmt(total)} FCFA
        </div>
        <button style={{ ...S.btn, ...S.btnSecondary }} onClick={() => generateBordereauExcel(factures)}>
          📥 Générer Bordereau Excel
        </button>
      </div>

      {loading ? (
        <div style={{ textAlign: 'center', padding: 40, color: '#64748B' }}>Chargement...</div>
      ) : (
        <div style={{ overflowX: 'auto' }}>
          <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 13 }}>
            <thead>
              <tr>
                {['N° Facture','Date','Compagnie','Site','Période','H. Comptoirs','PAX','Montant HT','Statut','Actions'].map(h => (
                  <th key={h} style={S.tableHeader}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {factures.length === 0 ? (
                <tr><td colSpan={10} style={{ ...S.tableCell, textAlign: 'center', color: '#94A3B8', padding: 32 }}>Aucune facture</td></tr>
              ) : factures.map((f, i) => (
                <tr key={f.id} style={{ background: i % 2 === 0 ? '#fff' : '#F8FAFC' }}>
                  <td style={{ ...S.tableCell, fontWeight: 700, color: '#1A2B4A' }}>{f.numero_facture}</td>
                  <td style={S.tableCell}>{new Date(f.date_facture).toLocaleDateString('fr-FR')}</td>
                  <td style={S.tableCell}>{f.compagnie}</td>
                  <td style={S.tableCell}>{f.site}</td>
                  <td style={S.tableCell}>
                    {f.periode_debut && f.periode_fin ? `${new Date(f.periode_debut).toLocaleDateString('fr-FR')} → ${new Date(f.periode_fin).toLocaleDateString('fr-FR')}` : '—'}
                  </td>
                  <td style={S.tableCell}>{f.nombre_heures}h</td>
                  <td style={S.tableCell}>{f.total_pax}</td>
                  <td style={{ ...S.tableCell, fontWeight: 700 }}>{fmt(f.montant_ht)} FCFA</td>
                  <td style={S.tableCell}><span style={S.badge(f.statut)}>{f.statut}</span></td>
                  <td style={S.tableCell}>
                    <div style={{ display: 'flex', gap: 4 }}>
                      <button style={{ ...S.btn, padding: '4px 10px', fontSize: 11, background: '#EFF6FF', color: '#1E40AF', border: '1px solid #BFDBFE' }}
                        onClick={() => generateFactureExcel(f)}>Excel</button>
                      {f.statut === 'brouillon' && <button style={{ ...S.btn, ...S.btnGreen, padding: '4px 8px', fontSize: 11 }} onClick={() => handleStatusChange(f, 'emise')}>Émettre</button>}
                      {f.statut === 'emise' && <button style={{ ...S.btn, padding: '4px 8px', fontSize: 11, background: '#D1FAE5', color: '#065F46', border: '1px solid #A7F3D0' }} onClick={() => handleStatusChange(f, 'payee')}>Payée</button>}
                      <button style={{ ...S.btn, ...S.btnDanger, padding: '4px 8px', fontSize: 11 }} onClick={() => handleDelete(f.id)}>✕</button>
                    </div>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}

// ── Main BandesModule ─────────────────────────────────────────────────────────

export function BandesModule() {
  const [subTab, setSubTab] = useState<'saisie' | 'fiches' | 'facturation' | 'bordereau'>('saisie');
  const [editFiche, setEditFiche] = useState<Fiche | null>(null);
  const [ficheIdsToFacture, setFicheIdsToFacture] = useState<string[]>([]);
  const [allFiches, setAllFiches] = useState<Fiche[]>([]);
  const [refreshKey, setRefreshKey] = useState(0);
  const [globalToast, setGlobalToast] = useState('');

  const loadFiches = useCallback(async () => {
    const data = await apiFetch('GET', '/fiches-bandes');
    if (data.success) setAllFiches(data.data);
  }, []);

  useEffect(() => { loadFiches(); }, [loadFiches, refreshKey]);

  const handleSaved = () => {
    setRefreshKey(k => k + 1);
    setEditFiche(null);
    if (subTab === 'saisie') setSubTab('fiches');
  };

  const handleEdit = (f: Fiche) => { setEditFiche(f); setSubTab('saisie'); };

  const handleFacturer = (ids: string[]) => {
    setFicheIdsToFacture(ids);
    setSubTab('facturation');
  };

  const handleFactureSaved = () => {
    setRefreshKey(k => k + 1);
    setFicheIdsToFacture([]);
    setSubTab('bordereau');
  };

  const tabs = [
    { key: 'saisie', label: editFiche ? 'Modifier fiche' : 'Saisie fiche' },
    { key: 'fiches', label: 'Fiches' },
    { key: 'facturation', label: 'Facturation' },
    { key: 'bordereau', label: 'Bordereau' },
  ] as const;

  return (
    <div>
      {globalToast && <Toast msg={globalToast} onClose={() => setGlobalToast('')} />}
      {/* Sub-tab nav */}
      <div style={{ borderBottom: '2px solid #E2E8F0', display: 'flex', gap: 0 }}>
        {tabs.map(t => (
          <button key={t.key}
            onClick={() => { setSubTab(t.key); if (t.key !== 'saisie') setEditFiche(null); }}
            style={{
              padding: '12px 24px', fontSize: 14, fontWeight: 600, border: 'none',
              borderBottom: subTab === t.key ? '2px solid #1A2B4A' : '2px solid transparent',
              marginBottom: -2, cursor: 'pointer', background: 'transparent',
              color: subTab === t.key ? '#1A2B4A' : '#64748B',
              fontFamily: 'Inter, sans-serif', transition: 'all 0.2s',
            }}>
            {t.label}
          </button>
        ))}
      </div>

      {subTab === 'saisie' && <SaisieForm onSaved={handleSaved} editFiche={editFiche} onToast={setGlobalToast} />}
      {subTab === 'fiches' && <FichesList onEdit={handleEdit} onFacturer={handleFacturer} onRefresh={refreshKey} />}
      {subTab === 'facturation' && <FacturationForm ficheIds={ficheIdsToFacture} fiches={allFiches} onSaved={handleFactureSaved} />}
      {subTab === 'bordereau' && <Bordereau />}
    </div>
  );
}
