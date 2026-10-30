/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useEffect } from 'react';
import { Plus, Trash2, Printer, Download, Info, RotateCcw, FileDown, ChevronDown, ChevronUp, ExternalLink } from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import * as XLSX from 'xlsx';
import { clsx, type ClassValue } from 'clsx';
import { twMerge } from 'tailwind-merge';

function cn(...inputs: ClassValue[]) {
  return twMerge(clsx(inputs));
}

interface Sequence {
  id: string;
  competence: string;
  objectif: string;
  contenu: string;
  theme: string;
  duree: number;
  modalite: string;
  methode: string;
  methodeAutre: string;
  activite: string;
  outil: string;
  evaluation: string;
}

const STORAGE_KEY = 'cce_formateur_lms_save';
const TARGET_MINUTES = 420;

const MODALITIES = ["Présentiel", "Présentiel enrichi", "Distanciel synchrone", "Distanciel asynchrone", "Hybride"];
const METHODS = ["Expositive", "Démonstrative", "Interrogative", "Active", "Expérientielle", "Autre"];

const INITIAL_EXAMPLE: Sequence = {
  id: 'example',
  competence: "Mobiliser la communication non violente en situation conflictuelle",
  objectif: "Appliquer les méthodes de la communication non violente lors d'une mise en situation",
  contenu: "Les 4 étapes de la communication non violente (OSBD : Observation, Sentiment, Besoin, Demande)",
  theme: "Jeu de rôle – Client mécontent",
  duree: 45,
  modalite: "Présentiel",
  methode: "Expérientielle",
  methodeAutre: "",
  activite: "Jeu de rôle en binôme avec débriefing collectif",
  outil: "Fiches scénarios, grille d'observation",
  evaluation: "Auto-évaluation guidée"
};

export default function App() {
  const [sequences, setSequences] = useState<Sequence[]>([]);
  const [isLoaded, setIsLoaded] = useState(false);
  const [isAccordionOpen, setIsAccordionOpen] = useState(true);

  const createNewSequence = (): Sequence => ({
    id: Math.random().toString(36).substr(2, 9),
    competence: "",
    objectif: "",
    contenu: "",
    theme: "",
    duree: 0,
    modalite: MODALITIES[0],
    methode: METHODS[0],
    methodeAutre: "",
    activite: "",
    outil: "",
    evaluation: ""
  });

  // Load data
  useEffect(() => {
    const saved = localStorage.getItem(STORAGE_KEY);
    if (saved) {
      try {
        setSequences(JSON.parse(saved));
      } catch (e) {
        console.error("Failed to load sequences", e);
        setSequences(Array.from({ length: 5 }, () => createNewSequence()));
      }
    } else {
      // Default to 5 empty rows
      setSequences(Array.from({ length: 5 }, () => createNewSequence()));
    }
    setIsLoaded(true);
  }, []);

  // Save data
  useEffect(() => {
    if (isLoaded) {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(sequences));
    }
  }, [sequences, isLoaded]);

  const addRow = () => {
    setSequences(prev => [...prev, createNewSequence()]);
  };

  const removeRow = (id: string) => {
    setSequences(prev => prev.filter(s => s.id !== id));
  };

  const updateRow = (id: string, field: keyof Sequence, value: string | number) => {
    setSequences(prev => prev.map(s => s.id === id ? { ...s, [field]: value } : s));
  };

  const resetData = () => {
    if (confirm("Effacer tout votre travail ?")) {
      setSequences(Array.from({ length: 5 }, () => createNewSequence()));
    }
  };

  const totalMinutes = sequences.reduce((acc, s) => acc + (Number(s.duree) || 0), 0);

  const exportXls = (isTpl = false) => {
    if (!isTpl && totalMinutes !== TARGET_MINUTES) {
      if (!confirm(`Attention : Votre total est de ${totalMinutes} min. Souhaitez-vous quand même exporter ?`)) return;
    }

    const head = ["Compétence", "Objectif", "Contenu", "Thème", "Durée", "Modalité", "Méthode", "Activité", "Outil", "Évaluation"];
    const data: (string | number)[][] = [head];

    data.push([
      INITIAL_EXAMPLE.competence,
      INITIAL_EXAMPLE.objectif,
      INITIAL_EXAMPLE.contenu,
      INITIAL_EXAMPLE.theme,
      INITIAL_EXAMPLE.duree,
      INITIAL_EXAMPLE.modalite,
      INITIAL_EXAMPLE.methode,
      INITIAL_EXAMPLE.activite,
      INITIAL_EXAMPLE.outil,
      INITIAL_EXAMPLE.evaluation
    ]);

    if (!isTpl) {
      sequences.forEach(s => {
        const me = s.methode;
        data.push([
          s.competence,
          s.objectif,
          s.contenu,
          s.theme,
          s.duree,
          s.modalite,
          me === 'Autre' ? s.methodeAutre : me,
          s.activite,
          s.outil,
          s.evaluation
        ]);
      });
    }

    const ws = XLSX.utils.aoa_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Déroulé");
    XLSX.writeFile(wb, isTpl ? "Template_CCE.xlsx" : "Mon_Deroule_Pedagogique.xlsx");

    if (!isTpl) {
      setTimeout(() => {
        window.open("https://1001parcours.cci.fr/experience/mod/page/view.php?id=2092", "_blank");
      }, 1500);
    }
  };

  return (
    <div className="min-h-screen p-4 md:p-8">
      <div className="container mx-auto bg-white rounded-2xl shadow-xl p-6 md:p-8">
        {/* Main Title */}
        <h1 className="text-3xl font-bold text-center text-gray-900 mb-8 no-print">Rédiger un déroulé pédagogique</h1>

        {/* Header Section */}
        <header className="flex flex-col md:flex-row justify-between items-start gap-6 mb-8 no-print">
          <div className="flex flex-col gap-3 w-full md:w-auto order-2 md:order-1">
            <a 
              href="https://1001parcours.cci.fr/experience/mod/page/view.php?id=2092"
              className="flex items-center justify-center gap-2 px-4 py-2 bg-gray-100 text-gray-700 font-semibold rounded-lg hover:bg-gray-200 transition-colors border border-gray-200"
            >
              <ExternalLink size={18} />
              Retour au parcours
            </a>
            <button 
              onClick={() => exportXls(true)}
              className="flex items-center justify-center gap-2 px-4 py-2 border-2 border-primary text-primary font-semibold rounded-lg hover:bg-header-bg transition-colors"
            >
              <FileDown size={18} />
              Template Excel
            </button>
            <button 
              onClick={() => window.print()}
              className="flex items-center justify-center gap-2 px-4 py-2 border-2 border-primary text-primary font-semibold rounded-lg hover:bg-header-bg transition-colors"
            >
              <Printer size={18} />
              Imprimer / PDF
            </button>
          </div>

          <div className="flex-1 w-full order-1 md:order-2">
            {/* Accordion Consigne */}
            <div className="border border-primary/20 rounded-xl overflow-hidden shadow-sm">
              <button 
                onClick={() => setIsAccordionOpen(!isAccordionOpen)}
                className="w-full flex items-center justify-between p-4 bg-header-bg hover:bg-primary/10 transition-colors text-primary font-bold text-lg"
              >
                <div className="flex items-center gap-2">
                  <Info size={20} />
                  Consigne
                </div>
                {isAccordionOpen ? <ChevronUp size={20} /> : <ChevronDown size={20} />}
              </button>
              
              <AnimatePresence>
                {isAccordionOpen && (
                  <motion.div 
                    initial={{ height: 0, opacity: 0 }}
                    animate={{ height: 'auto', opacity: 1 }}
                    exit={{ height: 0, opacity: 0 }}
                    transition={{ duration: 0.3, ease: 'easeInOut' }}
                    className="overflow-hidden"
                  >
                    <div className="p-6 bg-white border-t border-primary/10">
                      <p className="text-gray-800 font-medium mb-4">
                        Concevez le déroulé pédagogique d'une journée de formation de 7 heures sur une thématique de votre choix ou à partir du cas Résid'Azur.
                      </p>
                      <div className="grid md:grid-cols-2 gap-6">
                        <div>
                          <h3 className="text-primary font-semibold mb-2">Votre déroulé doit inclure :</h3>
                          <ul className="text-sm space-y-1 list-disc pl-5 text-gray-700">
                            <li>Un accueil des participants</li>
                            <li>Un brise-glace</li>
                            <li>Une évaluation des acquis</li>
                            <li>Une évaluation à chaud en fin de journée</li>
                          </ul>
                        </div>
                        <div>
                          <h3 className="text-primary font-semibold mb-2">Conseils :</h3>
                          <ul className="text-sm space-y-1 list-disc pl-5 text-gray-700">
                            <li>Complétez toutes les colonnes (indiquez NA si non applicable)</li>
                            <li>Surveillez le compteur pour atteindre 420 minutes</li>
                            <li>Survolez les en-têtes pour afficher les infobulles d'aide</li>
                            <li>Une ligne d'exemple est fournie à titre de référence.</li>
                          </ul>
                        </div>
                      </div>
                    </div>
                  </motion.div>
                )}
              </AnimatePresence>
            </div>

            {/* Add Row Button - Moved here */}
            <div className="flex justify-end mt-4 no-print">
              <button 
                onClick={addRow}
                className="flex items-center justify-center gap-2 px-6 py-2 border-2 border-primary text-primary font-bold rounded-xl hover:bg-header-bg transition-all hover:scale-[1.02] active:scale-[0.98]"
              >
                <Plus size={20} />
                Ajouter une ligne
              </button>
            </div>
          </div>
        </header>

        {/* Table Section */}
        <div className="overflow-x-auto border border-gray-200 rounded-xl shadow-sm mb-8">
          <table className="w-full border-collapse min-w-[1400px] table-fixed">
            <thead className="sticky top-0 z-10">
              <tr className="bg-header-bg border-b-2 border-gray-200">
                <HeaderCell title="Compétence visée" tooltip="Indiquez la compétence professionnelle que cette séquence contribue à développer, en lien avec le référentiel de formation ou les attendus métier." width="11%" />
                <HeaderCell title="Objectif pédagogique" tooltip="Formulez l'objectif avec un verbe d'action observable à l'aide de la taxonomie de Bloom (ex : identifier, appliquer, analyser). Précisez ce que l'apprenant sera capable de faire à l'issue de la séquence." width="11%" />
                <HeaderCell title="Contenu" tooltip="Décrivez les notions, concepts ou savoir-faire abordés durant cette séquence." width="11%" />
                <HeaderCell title="Thème de la séquence" tooltip="Donnez un titre court et explicite à cette séquence (ex : 'Accueil et présentation', 'Mise en situation client difficile')." width="10%" />
                <HeaderCell title="Durée" tooltip="Indiquez la durée en minutes. Le total de la journée doit atteindre 420 minutes (7 heures)." width="80px" />
                <HeaderCell title="Modalité" tooltip="Sélectionnez le format de la séquence : présentiel, distanciel, hybride, etc." width="10%" />
                <HeaderCell title="Méthode pédagogique" tooltip="Choisissez la méthode dominante : expositive (transmission), démonstrative (montrer), interrogative (questionnement), active (mise en action), expérientielle (apprentissage par l'expérience)." width="10%" />
                <HeaderCell title="Activité pédagogique" tooltip="Décrivez concrètement ce que font les apprenants et le formateur (ex : brainstorming, étude de cas, jeu de rôle, quiz)." width="11%" />
                <HeaderCell title="Outil" tooltip="Listez les supports et outils utilisés (ex : diaporama, vidéo, paperboard, application numérique). Indiquez NA si non applicable." width="10%" />
                <HeaderCell title="Évaluation" tooltip="Précisez comment vous vérifiez l'atteinte de l'objectif (ex : QCM, observation, auto-évaluation). Indiquez NA si non applicable." width="10%" />
                <th className="w-[50px] no-print"></th>
              </tr>
            </thead>
            <tbody>
              {/* Example Row */}
              <tr className="bg-header-bg/30 italic text-gray-500 text-sm">
                <td className="p-3 border-b border-gray-100">{INITIAL_EXAMPLE.competence}</td>
                <td className="p-3 border-b border-gray-100">{INITIAL_EXAMPLE.objectif}</td>
                <td className="p-3 border-b border-gray-100">{INITIAL_EXAMPLE.contenu}</td>
                <td className="p-3 border-b border-gray-100">{INITIAL_EXAMPLE.theme}</td>
                <td className="p-3 border-b border-gray-100">{INITIAL_EXAMPLE.duree}</td>
                <td className="p-3 border-b border-gray-100">{INITIAL_EXAMPLE.modalite}</td>
                <td className="p-3 border-b border-gray-100">{INITIAL_EXAMPLE.methode}</td>
                <td className="p-3 border-b border-gray-100">{INITIAL_EXAMPLE.activite}</td>
                <td className="p-3 border-b border-gray-100">{INITIAL_EXAMPLE.outil}</td>
                <td className="p-3 border-b border-gray-100">{INITIAL_EXAMPLE.evaluation}</td>
                <td className="p-3 border-b border-gray-100 no-print text-xs font-semibold">Ex.</td>
              </tr>

              {/* Dynamic Rows */}
              <AnimatePresence initial={false}>
                {sequences.map((seq) => (
                  <motion.tr 
                    key={seq.id}
                    initial={{ opacity: 0, y: 10 }}
                    animate={{ opacity: 1, y: 0 }}
                    exit={{ opacity: 0, scale: 0.95 }}
                    className="group hover:bg-gray-50 transition-colors"
                  >
                    <Cell><Textarea value={seq.competence} onChange={v => updateRow(seq.id, 'competence', v)} /></Cell>
                    <Cell><Textarea value={seq.objectif} onChange={v => updateRow(seq.id, 'objectif', v)} /></Cell>
                    <Cell><Textarea value={seq.contenu} onChange={v => updateRow(seq.id, 'contenu', v)} /></Cell>
                    <Cell><Textarea value={seq.theme} onChange={v => updateRow(seq.id, 'theme', v)} /></Cell>
                    <Cell>
                      <input 
                        type="number" 
                        value={seq.duree} 
                        onChange={e => updateRow(seq.id, 'duree', parseInt(e.target.value) || 0)}
                        className="w-full p-2 border border-gray-300 rounded focus:ring-2 focus:ring-primary focus:border-transparent outline-none text-sm"
                        min="0"
                      />
                    </Cell>
                    <Cell>
                      <Select 
                        value={seq.modalite} 
                        options={MODALITIES} 
                        onChange={v => updateRow(seq.id, 'modalite', v)} 
                      />
                    </Cell>
                    <Cell>
                      <Select 
                        value={seq.methode} 
                        options={METHODS} 
                        onChange={v => updateRow(seq.id, 'methode', v)} 
                      />
                      {seq.methode === 'Autre' && (
                        <input 
                          type="text"
                          placeholder="Précisez..."
                          value={seq.methodeAutre}
                          onChange={e => updateRow(seq.id, 'methodeAutre', e.target.value)}
                          className="w-full mt-2 p-2 border border-gray-300 rounded text-xs"
                        />
                      )}
                    </Cell>
                    <Cell><Textarea value={seq.activite} onChange={v => updateRow(seq.id, 'activite', v)} /></Cell>
                    <Cell><Textarea value={seq.outil} onChange={v => updateRow(seq.id, 'outil', v)} /></Cell>
                    <Cell><Textarea value={seq.evaluation} onChange={v => updateRow(seq.id, 'evaluation', v)} /></Cell>
                    <td className="p-3 border-b border-gray-100 no-print align-middle">
                      <button 
                        onClick={() => removeRow(seq.id)}
                        className="p-2 text-red-500 hover:bg-red-50 rounded-lg transition-colors opacity-0 group-hover:opacity-100"
                        title="Supprimer"
                      >
                        <Trash2 size={18} />
                      </button>
                    </td>
                  </motion.tr>
                ))}
              </AnimatePresence>
            </tbody>
          </table>
        </div>

        {/* Counter Section */}
        <div 
          className={cn(
            "mb-8 p-3 rounded-xl text-center text-base font-bold transition-all duration-300 shadow-inner border-2",
            totalMinutes === TARGET_MINUTES 
              ? "bg-green-100 text-green-700 border-green-500" 
              : "bg-amber-100 text-amber-700 border-amber-500"
          )}
        >
          Total : {totalMinutes} min / {TARGET_MINUTES} min (7h)
        </div>

        {/* Footer Actions */}
        <footer className="flex flex-col md:flex-row justify-between items-end gap-6 no-print">
          <div className="flex flex-col gap-4 w-full md:w-auto">
            <button 
              onClick={resetData}
              className="flex items-center justify-center gap-2 px-4 py-2 text-red-500 hover:text-red-700 font-medium transition-colors"
            >
              <RotateCcw size={16} />
              Réinitialiser tout
            </button>
          </div>
          
          <button 
            onClick={() => exportXls(false)}
            className="w-full md:w-auto flex items-center justify-center gap-3 px-8 py-4 bg-primary text-white font-bold rounded-xl hover:bg-primary-hover shadow-lg hover:shadow-primary/30 transition-all hover:scale-[1.02] active:scale-[0.98]"
          >
            <Download size={22} />
            Valider et Télécharger (.xlsx)
          </button>
        </footer>
      </div>
    </div>
  );
}

function HeaderCell({ title, tooltip, width }: { title: string, tooltip: string, width: string }) {
  return (
    <th className="p-4 text-left text-xs font-bold uppercase tracking-wider text-gray-700 relative group" style={{ width }}>
      <div className="flex items-center gap-1">
        <span className="whitespace-normal leading-tight">{title}</span>
        <div className="cursor-help p-0.5 bg-primary text-white rounded-full flex-shrink-0">
          <Info size={10} />
        </div>
        {/* Tooltip - Positioned below to avoid clipping by table container top */}
        <div className="invisible opacity-0 group-hover:visible group-hover:opacity-100 absolute z-50 w-64 p-3 bg-gray-800 text-white text-xs font-normal normal-case rounded-lg shadow-xl transition-all duration-200 pointer-events-none -translate-x-1/2 left-1/2 top-full mt-2 border border-gray-700">
          {tooltip}
          <div className="absolute bottom-full left-1/2 -translate-x-1/2 border-8 border-transparent border-b-gray-800"></div>
        </div>
      </div>
    </th>
  );
}

function Cell({ children }: { children: React.ReactNode }) {
  return <td className="p-3 border-b border-gray-100 align-top">{children}</td>;
}

function Textarea({ value, onChange }: { value: string, onChange: (v: string) => void }) {
  return (
    <textarea 
      value={value}
      onChange={e => onChange(e.target.value)}
      className="w-full p-2 border border-gray-300 rounded focus:ring-2 focus:ring-primary focus:border-transparent outline-none text-sm min-h-[60px] resize-y"
    />
  );
}

function Select({ value, options, onChange }: { value: string, options: string[], onChange: (v: string) => void }) {
  return (
    <select 
      value={value}
      onChange={e => onChange(e.target.value)}
      className="w-full p-2 border border-gray-300 rounded focus:ring-2 focus:ring-primary focus:border-transparent outline-none text-sm bg-white"
    >
      {options.map(opt => (
        <option key={opt} value={opt}>{opt}</option>
      ))}
    </select>
  );
}
