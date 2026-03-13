import { useState } from "react";
import { motion, AnimatePresence } from "framer-motion";
import { ChevronRight, ChevronLeft, Building2, Landmark, Layers } from "lucide-react";

function Field({ label, type = "text", placeholder, options = [], value, onChange }) {
  const inputClass =
    "w-full rounded-xl border border-slate-200 px-4 py-3 text-sm outline-none focus:border-slate-400 focus:ring-2 focus:ring-slate-200";

  return (
    <label className="block">
      <span className="mb-2 block text-sm font-medium text-slate-700">{label}</span>
      {type === "select" ? (
        <select className={inputClass} value={value} onChange={(e) => onChange(e.target.value)}>
          {options.map((opt) => (
            <option key={opt}>{opt}</option>
          ))}
        </select>
      ) : (
        <input
          type={type}
          placeholder={placeholder}
          value={value}
          onChange={(e) => onChange(e.target.value)}
          className={inputClass}
        />
      )}
    </label>
  );
}

export default function CmaWizard() {
  const [step, setStep] = useState(1);

  const [businessType, setBusinessType] = useState("New Business");
  const [loanType, setLoanType] = useState("CC");

  const [clientName, setClientName] = useState("");
  const [firmName, setFirmName] = useState("");
  const [bankName, setBankName] = useState("");

  const showCC = loanType === "CC" || loanType === "CC + Term Loan";
  const showTL = loanType === "Term Loan" || loanType === "CC + Term Loan";

  const next = () => setStep((s) => Math.min(s + 1, 4));
  const back = () => setStep((s) => Math.max(s - 1, 1));

  return (
    <div className="min-h-screen bg-slate-50 p-6">
      <div className="mx-auto max-w-4xl">
        <div className="mb-8 rounded-2xl bg-white p-6 shadow">
          <h1 className="text-2xl font-bold text-slate-900">CMA Projection Wizard</h1>
          <p className="text-sm text-slate-500">Step based CMA input system</p>

          <div className="mt-6 flex gap-4 text-sm">
            {["Basic", "Loan", "Inputs", "Preview"].map((s, i) => (
              <div
                key={s}
                className={`flex items-center gap-2 rounded-full px-4 py-2 ${
                  step === i + 1 ? "bg-slate-900 text-white" : "bg-slate-200"
                }`}
              >
                {i + 1}. {s}
              </div>
            ))}
          </div>
        </div>

        <AnimatePresence mode="wait">
          {step === 1 && (
            <motion.div
              key="step1"
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              exit={{ opacity: 0 }}
              className="card"
            >
              <div className="rounded-2xl bg-white p-6 shadow">
                <div className="mb-4 flex items-center gap-2 text-lg font-semibold">
                  <Layers className="h-5 w-5" /> Basic Details
                </div>

                <div className="grid gap-4 md:grid-cols-2">
                  <Field label="Client Name" value={clientName} onChange={setClientName} />
                  <Field label="Firm Name" value={firmName} onChange={setFirmName} />
                  <Field label="Bank Name" value={bankName} onChange={setBankName} />
                </div>
              </div>
            </motion.div>
          )}

          {step === 2 && (
            <motion.div key="step2" initial={{ opacity: 0 }} animate={{ opacity: 1 }} exit={{ opacity: 0 }}>
              <div className="rounded-2xl bg-white p-6 shadow">
                <div className="mb-4 flex items-center gap-2 text-lg font-semibold">
                  <Building2 className="h-5 w-5" /> Loan Selection
                </div>

                <div className="grid gap-4 md:grid-cols-2">
                  <Field
                    label="Business Type"
                    type="select"
                    value={businessType}
                    onChange={setBusinessType}
                    options={["New Business", "Existing Business"]}
                  />

                  <Field
                    label="Loan Type"
                    type="select"
                    value={loanType}
                    onChange={setLoanType}
                    options={["CC", "Term Loan", "CC + Term Loan"]}
                  />
                </div>
              </div>
            </motion.div>
          )}

          {step === 3 && (
            <motion.div key="step3" initial={{ opacity: 0 }} animate={{ opacity: 1 }} exit={{ opacity: 0 }}>
              {showCC && (
                <div className="mb-6 rounded-2xl bg-white p-6 shadow">
                  <h2 className="mb-4 text-lg font-semibold">CC Details</h2>

                  <div className="grid gap-4 md:grid-cols-2">
                    <Field label="Proposed CC Limit" type="number" />
                    <Field label="ROI (%)" type="number" />
                    <Field label="Debtors Days" type="number" />
                    <Field label="Creditors Days" type="number" />
                  </div>
                </div>
              )}

              {showTL && (
                <div className="rounded-2xl bg-white p-6 shadow">
                  <h2 className="mb-4 flex items-center gap-2 text-lg font-semibold">
                    <Landmark className="h-5 w-5" /> Term Loan Details
                  </h2>

                  <div className="grid gap-4 md:grid-cols-2">
                    <Field label="Term Loan Amount" type="number" />
                    <Field label="ROI (%)" type="number" />
                    <Field label="Repayment Months" type="number" />
                    <Field label="Moratorium" type="number" />
                  </div>
                </div>
              )}

              {businessType === "Existing Business" && (
                <div className="mt-6 rounded-2xl bg-white p-6 shadow">
                  <h2 className="mb-4 text-lg font-semibold">Existing Business Inputs</h2>

                  <div className="grid gap-4 md:grid-cols-2">
                    <Field label="Last Year Sales" type="number" />
                    <Field label="Existing Net Worth" type="number" />
                    <Field label="Closing Stock" type="number" />
                  </div>
                </div>
              )}
            </motion.div>
          )}

          {step === 4 && (
            <motion.div key="step4" initial={{ opacity: 0 }} animate={{ opacity: 1 }} exit={{ opacity: 0 }}>
              <div className="rounded-2xl bg-white p-6 shadow">
                <h2 className="mb-4 text-lg font-semibold">Preview</h2>

                <p className="text-sm text-slate-600">Client: {clientName}</p>
                <p className="text-sm text-slate-600">Firm: {firmName}</p>
                <p className="text-sm text-slate-600">Loan Type: {loanType}</p>
                <p className="text-sm text-slate-600">Business Type: {businessType}</p>
              </div>
            </motion.div>
          )}
        </AnimatePresence>

        <div className="mt-6 flex justify-between">
          <button onClick={back} className="flex items-center gap-2 rounded-xl border px-4 py-2 text-sm">
            <ChevronLeft size={16} /> Back
          </button>

          <button
            onClick={next}
            className="flex items-center gap-2 rounded-xl bg-slate-900 px-4 py-2 text-sm text-white"
          >
            Next <ChevronRight size={16} />
          </button>
        </div>
      </div>
    </div>
  );
}
