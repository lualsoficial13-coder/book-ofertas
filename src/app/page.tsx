"use client";

import {
  FormEvent,
  useCallback,
  useEffect,
  useMemo,
  useState,
} from "react";
import * as XLSX from "xlsx";
import { jsPDF } from "jspdf";
import { supabase } from "./lib/supabase";

type HistoryItem = {
  action: string;
  person: string;
  dateTime: string;
};

type Offer = {
  id: string;
  code: string;
  product: string;
  crop: string;
  origin: string;
  supplier: string;
  shippingLocation: string;
  destination: string;
  commercial: string;
  quantityKg: number;
  askingPrice: number;
  ourPrice: number;
  withdrawalType: string;
  quality: string;
  damagedPercentage: string;
  status: string;
  notes: string;
  createdBy: string;
  createdById: string;
  createdAt: string;
  createdAtIso: string;
  history: HistoryItem[];
};

type OfferRow = {
  id: string;
  code: string;
  product: string;
  crop: string;
  origin: string;
  supplier: string;
  shipping_location: string;
  destination: string;
  commercial: string;
  quantity_kg: number;
  asking_price: number;
  our_price: number;
  withdrawal_type: string;
  quality: string;
  damaged_percentage: string | null;
  status: string;
  notes: string | null;
  created_by: string;
  created_at: string;
};

type BoletaData = Record<string, string>;

type BoletaRow = {
  id: string;
  offer_id: string;
  data: BoletaData;
  updated_at?: string;
};

type HistoryRow = {
  id: string;
  offer_id: string;
  action: string;
  actor_id: string;
  actor_email: string;
  created_at: string;
};

type TextFieldProps = {
  label: string;
  value: string;
  onChange: (value: string) => void;
  placeholder?: string;
  type?: string;
  required?: boolean;
  min?: string;
  max?: string;
  step?: string;
};

function TextField({
  label,
  value,
  onChange,
  placeholder = "",
  type = "text",
  required = false,
  min,
  max,
  step,
}: TextFieldProps) {
  return (
    <div>
      <label className="mb-2 block text-sm font-medium text-slate-700">
        {label}
      </label>

      <input
        type={type}
        required={required}
        min={min}
        max={max}
        step={step}
        value={value}
        onChange={(event) => onChange(event.target.value)}
        placeholder={placeholder}
        className="w-full rounded-lg border border-slate-300 px-4 py-3 text-slate-900 outline-none focus:border-emerald-600 focus:ring-2 focus:ring-emerald-200"
      />
    </div>
  );
}

type SelectFieldProps = {
  label: string;
  value: string;
  onChange: (value: string) => void;
  options: string[];
};

function SelectField({
  label,
  value,
  onChange,
  options,
}: SelectFieldProps) {
  return (
    <div>
      <label className="mb-2 block text-sm font-medium text-slate-700">
        {label}
      </label>

      <select
        value={value}
        onChange={(event) => onChange(event.target.value)}
        className="w-full rounded-lg border border-slate-300 px-4 py-3 text-slate-900"
      >
        {options.map((option) => (
          <option key={option}>{option}</option>
        ))}
      </select>
    </div>
  );
}

const demoOffers: Offer[] = [];

function formatDateTime(iso: string) {
  if (!iso) return "";
  try {
    return new Date(iso).toLocaleString("pt-BR");
  } catch {
    return iso;
  }
}

function numberFormat(value: number) {
  return new Intl.NumberFormat("pt-BR").format(value || 0);
}

function money(value: number) {
  return new Intl.NumberFormat("pt-BR", {
    style: "currency",
    currency: "BRL",
  }).format(value || 0);
}

function statusColor(status: string) {
  if (status === "Fechada") return "bg-emerald-100 text-emerald-800";
  if (status === "Perdida" || status === "Cancelada") {
    return "bg-red-100 text-red-800";
  }
  if (status === "Enviada") return "bg-blue-100 text-blue-800";
  return "bg-amber-100 text-amber-800";
}

function rowToOffer(row: OfferRow, history: HistoryItem[]): Offer {
  return {
    id: row.id,
    code: row.code,
    product: row.product,
    crop: row.crop,
    origin: row.origin,
    supplier: row.supplier,
    shippingLocation: row.shipping_location,
    destination: row.destination,
    commercial: row.commercial,
    quantityKg: Number(row.quantity_kg) || 0,
    askingPrice: Number(row.asking_price) || 0,
    ourPrice: Number(row.our_price) || 0,
    withdrawalType: row.withdrawal_type,
    quality: row.quality,
    damagedPercentage: row.damaged_percentage || "",
    status: row.status,
    notes: row.notes || "",
    createdBy: row.created_by,
    createdById: row.created_by,
    createdAt: formatDateTime(row.created_at),
    createdAtIso: row.created_at,
    history,
  };
}

function historyRowToItem(row: HistoryRow): HistoryItem {
  return {
    action: row.action,
    person: row.actor_email || "Sistema",
    dateTime: formatDateTime(row.created_at),
  };
}

const BOLETA_FIELDS = [
  "buyer",
  "buyerDocument",
  "sellerDocument",
  "negotiationType",
  "paymentTerms",
  "paymentDate",
  "paymentMethod",
  "commissionOneName",
  "commissionOneValue",
  "commissionTwoName",
  "commissionTwoValue",
  "sellerStateRegistration",
  "sellerCityState",
  "contactPerson",
  "contactPhone",
  "contractEmail",
  "taxResponsible",
  "pisCofins",
  "funrural",
  "otherWithholdings",
  "commissionOneDocument",
  "commissionTwoDocument",
  "invoiceDeadline",
  "bank",
  "bankBranch",
  "bankAccount",
  "pixKey",
  "advancePayment",
  "advancePaymentDate",
  "advanceGuarantee",
  "accountOwnershipConfirmed",
  "pledgeCertificate",
  "sintegraChecked",
  "carValid",
  "shippingCityState",
  "shippingStart",
  "shippingEnd",
  "dailyTons",
] as const;

type BoletaState = {
  [K in (typeof BOLETA_FIELDS)[number]]: string;
};

function boletaDataAsRecord(): BoletaData {
  return BOLETA_FIELDS.reduce((acc, field) => {
    acc[field] = "";
    return acc;
  }, {} as BoletaData);
}

const defaultBoleta: BoletaState = {
  buyer: "Yellow Sorriso",
  buyerDocument: "",
  sellerDocument: "",
  negotiationType: "Compra no MI",
  paymentTerms: "Sobre Rodas",
  paymentDate: "",
  paymentMethod: "PIX",
  commissionOneName: "",
  commissionOneValue: "",
  commissionTwoName: "",
  commissionTwoValue: "",
  sellerStateRegistration: "",
  sellerCityState: "",
  contactPerson: "",
  contactPhone: "",
  contractEmail: "",
  taxResponsible: "Vendedor",
  pisCofins: "N/A",
  funrural: "N/A",
  otherWithholdings: "",
  commissionOneDocument: "",
  commissionTwoDocument: "",
  invoiceDeadline: "",
  bank: "",
  bankBranch: "",
  bankAccount: "",
  pixKey: "",
  advancePayment: "",
  advancePaymentDate: "",
  advanceGuarantee: "",
  accountOwnershipConfirmed: "Sim",
  pledgeCertificate: "N/A",
  sintegraChecked: "N/A",
  carValid: "N/A",
  shippingCityState: "",
  shippingStart: "",
  shippingEnd: "",
  dailyTons: "",
};

export default function Home() {
  const [authReady, setAuthReady] = useState(false);
  const [userEmail, setUserEmail] = useState("");
  const [userId, setUserId] = useState("");
  const [authMode, setAuthMode] = useState<"signin" | "signup" | "reset">(
    "signin"
  );
  const [authEmail, setAuthEmail] = useState("");
  const [authPassword, setAuthPassword] = useState("");
  const [authBusy, setAuthBusy] = useState(false);
  const [authMessage, setAuthMessage] = useState("");
  const [authError, setAuthError] = useState("");

  const [dataReady, setDataReady] = useState(false);
  const [dataError, setDataError] = useState("");
  const [offers, setOffers] = useState<Offer[]>(demoOffers);
  const [boletaRecords, setBoletaRecords] = useState<
    Record<string, BoletaData>
  >({});
  const [screen, setScreen] = useState<
    "dashboard" | "form" | "details" | "summary" | "boleta"
  >("dashboard");
  const [search, setSearch] = useState("");
  const [selectedOffer, setSelectedOffer] = useState<Offer | null>(null);
  const [editingId, setEditingId] = useState("");

  const [product, setProduct] = useState("Soja");
  const [crop, setCrop] = useState("2026/2027");
  const [origin, setOrigin] = useState("Norte MT");
  const [supplier, setSupplier] = useState("");
  const [shippingLocation, setShippingLocation] = useState("");
  const [destination, setDestination] = useState("");
  const [commercial, setCommercial] = useState("Antonio");
  const [quantityKg, setQuantityKg] = useState("");
  const [askingPrice, setAskingPrice] = useState("");
  const [ourPrice, setOurPrice] = useState("");
  const [withdrawalType, setWithdrawalType] = useState("FOB");
  const [quality, setQuality] = useState("Padrão");
  const [damagedPercentage, setDamagedPercentage] = useState("");
  const [status, setStatus] = useState("Em negociação");
  const [notes, setNotes] = useState("");

  const [boleta, setBoleta] = useState<BoletaState>({ ...defaultBoleta });
  const updateBoleta = useCallback(
    (field: keyof BoletaState, value: string) => {
      setBoleta((prev) => ({ ...prev, [field]: value }));
    },
    []
  );

  // ---------- Auth bootstrap ----------
  useEffect(() => {
    let mounted = true;
    async function bootstrap() {
      const { data } = await supabase.auth.getSession();
      if (!mounted) return;
      if (data.session?.user) {
        setUserEmail(data.session.user.email ?? "");
        setUserId(data.session.user.id ?? "");
      }
      setAuthReady(true);
    }
    bootstrap();
    const { data: sub } = supabase.auth.onAuthStateChange((_event, session) => {
      setUserEmail(session?.user?.email ?? "");
      setUserId(session?.user?.id ?? "");
      setAuthReady(true);
      if (!session) {
        setDataReady(false);
        setOffers([]);
        setBoletaRecords({});
        setSelectedOffer(null);
        setScreen("dashboard");
      }
    });
    return () => {
      mounted = false;
      sub.subscription.unsubscribe();
    };
  }, []);

  // ---------- Data load after login ----------
  useEffect(() => {
    if (!userId) return;
    let cancelled = false;
    (async () => {
      setDataReady(false);
      setDataError("");
      try {
        const [offersRes, boletasRes, historyRes] = await Promise.all([
          supabase
            .from("offers")
            .select("*")
            .order("created_at", { ascending: false }),
          supabase.from("boletas").select("*"),
          supabase
            .from("offer_history")
            .select("*")
            .order("created_at", { ascending: true }),
        ]);

        if (cancelled) return;

        if (offersRes.error) throw offersRes.error;
        if (boletasRes.error) throw boletasRes.error;
        if (historyRes.error) throw historyRes.error;

        const historyByOffer: Record<string, HistoryItem[]> = {};
        (historyRes.data || []).forEach((row: HistoryRow) => {
          if (!historyByOffer[row.offer_id]) {
            historyByOffer[row.offer_id] = [];
          }
          historyByOffer[row.offer_id].push(historyRowToItem(row));
        });

        const offerList: Offer[] = (offersRes.data || []).map(
          (row: OfferRow) =>
            rowToOffer(row, historyByOffer[row.id] || [])
        );

        const boletaMap: Record<string, BoletaData> = {};
        (boletasRes.data || []).forEach((row: BoletaRow) => {
          boletaMap[row.offer_id] = {
            ...boletaDataAsRecord(),
            ...(row.data || {}),
          };
        });

        setOffers(offerList);
        setBoletaRecords(boletaMap);
        setDataReady(true);
      } catch (err) {
        if (!cancelled) {
          const message =
            err instanceof Error
              ? err.message
              : "Não foi possível carregar os dados do servidor.";
          setDataError(message);
          setDataReady(true);
        }
      }
    })();
    return () => {
      cancelled = true;
    };
  }, [userId]);

  // ---------- Auth handlers ----------
  async function handleAuthSubmit(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    setAuthError("");
    setAuthMessage("");
    setAuthBusy(true);
    try {
      if (authMode === "signin") {
        const { error } = await supabase.auth.signInWithPassword({
          email: authEmail.trim(),
          password: authPassword,
        });
        if (error) throw error;
        setAuthMessage("Login realizado.");
      } else if (authMode === "signup") {
        const { error } = await supabase.auth.signUp({
          email: authEmail.trim(),
          password: authPassword,
        });
        if (error) throw error;
        setAuthMessage(
          "Conta criada. Verifique seu e-mail se a confirmação estiver habilitada."
        );
      } else if (authMode === "reset") {
        const redirectTo = `${window.location.origin}/reset-password`;
        const { error } = await supabase.auth.resetPasswordForEmail(
          authEmail.trim(),
          { redirectTo }
        );
        if (error) throw error;
        setAuthMessage(
          "Enviamos um link de recuperação para o seu e-mail. Abra-o em até 1 hora."
        );
      }
    } catch (err) {
      const message =
        err instanceof Error ? err.message : "Falha na operação de autenticação.";
      setAuthError(message);
    } finally {
      setAuthBusy(false);
    }
  }

  async function handleSignOut() {
    await supabase.auth.signOut();
  }

  // ---------- Offer helpers ----------
  async function recordHistory(offerId: string, action: string) {
    await supabase.from("offer_history").insert({
      offer_id: offerId,
      action,
      actor_id: userId,
      actor_email: userEmail,
    });
  }

  function clearOfferForm() {
    setEditingId("");
    setProduct("Soja");
    setCrop("2026/2027");
    setOrigin("Norte MT");
    setSupplier("");
    setShippingLocation("");
    setDestination("");
    setCommercial("Antonio");
    setQuantityKg("");
    setAskingPrice("");
    setOurPrice("");
    setWithdrawalType("FOB");
    setQuality("Padrão");
    setDamagedPercentage("");
    setStatus("Em negociação");
    setNotes("");
  }

  function openNewOffer() {
    clearOfferForm();
    setScreen("form");
  }

  async function openOffer(offer: Offer) {
    // track consultation in history (best effort)
    try {
      await recordHistory(offer.id, "Oferta aberta para consulta");
      setOffers((prev) =>
        prev.map((current) =>
          current.id === offer.id
            ? {
                ...current,
                history: [
                  ...current.history,
                  {
                    action: "Oferta aberta para consulta",
                    person: userEmail,
                    dateTime: new Date().toLocaleString("pt-BR"),
                  },
                ],
              }
            : current
        )
      );
    } catch {
      // ignore: still navigates
    }
    setSelectedOffer(offer);
    setScreen("details");
  }

  function editOffer(offer: Offer) {
    setEditingId(offer.id);
    setProduct(offer.product);
    setCrop(offer.crop);
    setOrigin(offer.origin);
    setSupplier(offer.supplier);
    setShippingLocation(offer.shippingLocation);
    setDestination(offer.destination);
    setCommercial(offer.commercial);
    setQuantityKg(String(offer.quantityKg));
    setAskingPrice(String(offer.askingPrice));
    setOurPrice(String(offer.ourPrice));
    setWithdrawalType(offer.withdrawalType);
    setQuality(offer.quality);
    setDamagedPercentage(offer.damagedPercentage);
    setStatus(offer.status);
    setNotes(offer.notes);
    setScreen("form");
  }

  async function saveOffer(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    setDataError("");

    const payload = {
      product,
      crop,
      origin,
      supplier,
      shipping_location: shippingLocation,
      destination,
      commercial,
      quantity_kg: Number(quantityKg),
      asking_price: Number(askingPrice),
      our_price: Number(ourPrice || 0),
      withdrawal_type: withdrawalType,
      quality,
      damaged_percentage: quality === "Avariado" ? damagedPercentage : null,
      status,
      notes: notes || "",
    };

    try {
      if (editingId) {
        const { data, error } = await supabase
          .from("offers")
          .update(payload)
          .eq("id", editingId)
          .select("*")
          .single();
        if (error) throw error;
        await recordHistory(editingId, "Oferta editada");
        const { data: hist } = await supabase
          .from("offer_history")
          .select("*")
          .eq("offer_id", editingId)
          .order("created_at", { ascending: true });

        const historyItems = (hist || []).map(historyRowToItem);
        const updated = rowToOffer(data as OfferRow, historyItems);
        setOffers((prev) =>
          prev.map((current) =>
            current.id === editingId ? updated : current
          )
        );
        setSelectedOffer(updated);
        clearOfferForm();
        setScreen("dashboard");
      } else {
        const { data, error } = await supabase
          .from("offers")
          .insert({ ...payload, created_by: userId })
          .select("*")
          .single();
        if (error) throw error;
        await recordHistory(data.id, "Oferta criada");
        const { data: hist } = await supabase
          .from("offer_history")
          .select("*")
          .eq("offer_id", data.id)
          .order("created_at", { ascending: true });

        const historyItems = (hist || []).map(historyRowToItem);
        const created = rowToOffer(data as OfferRow, historyItems);
        setOffers((prev) => [created, ...prev]);
        clearOfferForm();
        setScreen("dashboard");
      }
    } catch (err) {
      const message =
        err instanceof Error
          ? err.message
          : "Não foi possível salvar a oferta.";
      setDataError(message);
    }
  }

  function resetBoleta() {
    setBoleta({ ...defaultBoleta });
  }

  function loadBoletaFromRecord(record: BoletaData) {
    const next: BoletaState = { ...defaultBoleta };
    BOLETA_FIELDS.forEach((field) => {
      next[field] = record[field] ?? defaultBoleta[field];
    });
    setBoleta(next);
  }

  function openBoleta(offer: Offer) {
    setSelectedOffer(offer);
    const saved = boletaRecords[offer.id];
    if (saved) loadBoletaFromRecord(saved);
    else resetBoleta();
    setScreen("boleta");
  }

  async function persistBoleta() {
    if (!selectedOffer) return;
    const data: BoletaData = {
      ...boletaDataAsRecord(),
      ...boleta,
    };
    setDataError("");
    try {
      const { error } = await supabase
        .from("boletas")
        .upsert(
          {
            offer_id: selectedOffer.id,
            data,
            updated_at: new Date().toISOString(),
          },
          { onConflict: "offer_id" }
        );
      if (error) throw error;
      setBoletaRecords((prev) => ({
        ...prev,
        [selectedOffer.id]: data,
      }));
      await recordHistory(selectedOffer.id, "Boleta atualizada");
      setOffers((prev) =>
        prev.map((current) =>
          current.id === selectedOffer.id
            ? {
                ...current,
                history: [
                  ...current.history,
                  {
                    action: "Boleta atualizada",
                    person: userEmail,
                    dateTime: new Date().toLocaleString("pt-BR"),
                  },
                ],
              }
            : current
        )
      );
      window.alert("Dados da boleta salvos no servidor.");
    } catch (err) {
      const message =
        err instanceof Error
          ? err.message
          : "Não foi possível salvar a boleta.";
      setDataError(message);
    }
  }

  function getActiveBoleta(): BoletaData {
    if (!selectedOffer) return boletaDataAsRecord();
    const result: BoletaData = { ...boletaDataAsRecord(), ...boleta };
    return result;
  }

  function exportBackup() {
    const rows = offers.map((offer) => ({
      ...offer,
      history: JSON.stringify(offer.history),
      boletaData: JSON.stringify(boletaRecords[offer.id] || {}),
    }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(
      wb,
      XLSX.utils.json_to_sheet(rows),
      "Ofertas e Boletas"
    );
    XLSX.writeFile(wb, "backup-book-de-ofertas.xlsx");
  }

  function importBackupToLocal() {
    // Importação em massa exigirá endpoint específico — fluxo preservado apenas para Excel.
    window.alert(
      "A importação de backup externo só fica disponível após o usuário recriar as ofertas neste sistema."
    );
  }

  function downloadPdf() {
    if (!selectedOffer) return;
    const data = getActiveBoleta();
    const bags = selectedOffer.quantityKg / 60;
    const gross = bags * selectedOffer.ourPrice;
    const commission1 = bags * Number(data.commissionOneValue || 0);
    const commission2 = bags * Number(data.commissionTwoValue || 0);
    const daily = Number(data.dailyTons || 0);
    const estimatedDays = daily
      ? Math.ceil(selectedOffer.quantityKg / 1000 / daily)
      : 0;
    const document = new jsPDF({
      orientation: "portrait",
      unit: "mm",
      format: "a4",
    });
    const margin = 9,
      width = 192;
    let y = 9;
    const value = (item: string | number | undefined, fallback = "—") =>
      String(item || fallback).replace(/\s+/g, " ").trim();
    const field = (
      x: number,
      w: number,
      label: string,
      item: string | number | undefined,
      emphasis = false
    ) => {
      document.setDrawColor(148, 163, 184);
      document.setLineWidth(0.2);
      document.setFillColor(
        emphasis ? 226 : 255,
        emphasis ? 232 : 255,
        emphasis ? 240 : 255
      );
      document.rect(x, y, w, 7.7, "FD");
      document.setFont("helvetica", "bold");
      document.setFontSize(5.5);
      document.setTextColor(71, 85, 105);
      document.text(label.toUpperCase(), x + 1.6, y + 2.5);
      document.setFont("helvetica", emphasis ? "bold" : "normal");
      document.setFontSize(emphasis ? 7.3 : 6.8);
      document.setTextColor(15, 23, 42);
      document.text(
        document.splitTextToSize(value(item), w - 3).slice(0, 1),
        x + 1.6,
        y + 5.8
      );
    };
    const row = (
      a: string,
      av: string | number | undefined,
      b: string,
      bv: string | number | undefined,
      emphasis = false
    ) => {
      field(margin, width / 2, a, av, emphasis);
      field(margin + width / 2, width / 2, b, bv, emphasis);
      y += 7.7;
    };
    const section = (title: string) => {
      document.setFillColor(31, 78, 121);
      document.rect(margin, y, width, 4.3, "F");
      document.setFont("helvetica", "bold");
      document.setFontSize(6.8);
      document.setTextColor(255, 255, 255);
      document.text(title, margin + 1.7, y + 2.9);
      y += 4.3;
    };
    document.setFillColor(255, 255, 255);
    document.setDrawColor(31, 78, 121);
    document.setLineWidth(0.55);
    document.rect(margin, y, width, 14, "FD");
    document.setFont("helvetica", "bold");
    document.setFontSize(12);
    document.setTextColor(15, 23, 42);
    document.text(
      "BOLETA DE COMPRA / VENDA DE GRÃOS",
      margin + 3,
      y + 6.2
    );
    document.setFont("helvetica", "normal");
    document.setFontSize(6.3);
    document.setTextColor(71, 85, 105);
    document.text(
      "Preencha as informações e confira antes da assinatura.",
      margin + 3,
      y + 10
    );
    document.text("Código: FOR-COM-001  |  Revisão: 00", 145, y + 5);
    document.text("Página: 1 de 1", 169, y + 10);
    y += 15.5;
    section("1. IDENTIFICAÇÃO DA OPERAÇÃO");
    row("Nº da boleta", selectedOffer.code, "Data / hora do fechamento", selectedOffer.createdAt, true);
    row("Tipo de negociação", data.negotiationType, "Comercial responsável", selectedOffer.commercial);
    section("2. CONTRAPARTE");
    row("Comprador", data.buyer, "CNPJ do comprador", data.buyerDocument);
    row("Vendedor (razão social)", selectedOffer.supplier, "CNPJ / CPF do vendedor", data.sellerDocument);
    row("Inscrição estadual", data.sellerStateRegistration, "Município / UF", data.sellerCityState);
    row("Contato responsável", data.contactPerson, "Telefone", data.contactPhone);
    row("E-mail para assinatura do contrato", data.contractEmail, "Criada por", (selectedOffer.createdById === userId ? userEmail : "Membro da equipe"));
    section("3. PRODUTO E QUALIDADE");
    row("Produto", selectedOffer.product, "Safra", selectedOffer.crop);
    row("Classificação / padrão", selectedOffer.quality, "Avariado", selectedOffer.damagedPercentage ? `${selectedOffer.damagedPercentage}%` : "Não se aplica");
    section("4. QUANTIDADE E PREÇO");
    row("Quantidade contratada (sacas 60 kg)", numberFormat(bags), "Quantidade (kg)", numberFormat(selectedOffer.quantityKg), true);
    row("Preço unitário / saca (R$)", money(selectedOffer.ourPrice), "VALOR BRUTO DA MERCADORIA (R$)", money(gross), true);
    section("5. FRETE, COMISSÕES E TRIBUTOS");
    row("Modalidade do frete", selectedOffer.withdrawalType, "Imposto por conta de", data.taxResponsible);
    row("Operação com PIS / COFINS", data.pisCofins, "Funrural — folha ou produção", data.funrural);
    row("Outras retenções", data.otherWithholdings, "Base de cálculo da comissão", "Por saca (R$/sc)");
    row("Comissão 1", `${value(data.commissionOneName, "Não informado")} — ${value(data.commissionOneDocument)}`, "Valor / total", `${money(Number(data.commissionOneValue || 0))}/sc | ${money(commission1)}`);
    row("Comissão 2", `${value(data.commissionTwoName, "Não informado")} — ${value(data.commissionTwoDocument)}`, "Valor / total", `${money(Number(data.commissionTwoValue || 0))}/sc | ${money(commission2)}`);
    row("TOTAL DE COMISSÕES (R$)", money(commission1 + commission2), "", "", true);
    section("6. PAGAMENTO");
    row("Condição de pagamento", data.paymentTerms, "Data prevista de pagamento", data.paymentDate || "Não informada");
    row("Forma de pagamento", data.paymentMethod, "Prazo para emissão da NF", data.invoiceDeadline);
    row("Banco", data.bank, "Agência", data.bankBranch);
    row("Conta corrente", data.bankAccount, "Chave PIX", data.pixKey);
    row("Adiantamento (R$)", money(Number(data.advancePayment || 0)), "Data do adiantamento", data.advancePaymentDate || "Não informada");
    row("Garantia do adiantamento", data.advanceGuarantee, "Titularidade da conta confere?", data.accountOwnershipConfirmed);
    row("Certidão de penhor", data.pledgeCertificate, "Sintegra consultado", data.sintegraChecked);
    row("CAR válido", data.carValid, "", "");
    section("7. EMBARQUE E LOGÍSTICA");
    row("Local de embarque", selectedOffer.shippingLocation, "Município / UF de embarque", data.shippingCityState);
    row("Data início do embarque", data.shippingStart || "Não informada", "Data fim do embarque", data.shippingEnd || "Não informada");
    row("Cadência (ton/dia)", data.dailyTons ? `${data.dailyTons} ton/dia` : "Não informada", "Dias estimados de embarque", estimatedDays ? `${estimatedDays} dia(s)` : "—");
    document.setDrawColor(148, 163, 184);
    document.rect(margin, y, width, 9, "S");
    document.setFont("helvetica", "bold");
    document.setFontSize(5.5);
    document.setTextColor(71, 85, 105);
    document.text("OBSERVAÇÕES", margin + 1.6, y + 2.5);
    document.setFont("helvetica", "normal");
    document.setFontSize(6.5);
    document.setTextColor(15, 23, 42);
    document.text(
      document
        .splitTextToSize(value(selectedOffer.notes, "Nenhuma observação registrada."), width - 3)
        .slice(0, 2),
      margin + 1.6,
      y + 5.5
    );
    document.save(`boleta-${selectedOffer.code}.pdf`);
  }

  // ---------- Derived data ----------
  const totalVolume = useMemo(
    () => offers.reduce((total, offer) => total + (offer.quantityKg || 0), 0),
    [offers]
  );

  const visibleOffers = useMemo(() => {
    const term = search.trim().toLowerCase();
    if (!term) return offers;
    return offers.filter((offer) => offer.code.toLowerCase().includes(term));
  }, [offers, search]);

  const bags = Number(quantityKg || 0) / 60;

  const summaryOffers = useMemo(
    () =>
      offers.filter(
        (offer) => offer.status !== "Cancelada" && offer.status !== "Perdida"
      ),
    [offers]
  );

  const summaryVolume = useMemo(
    () => summaryOffers.reduce((t, o) => t + (o.quantityKg || 0), 0),
    [summaryOffers]
  );

  const summaryByMonth = useMemo(() => {
    const acc: Record<
      string,
      { key: string; label: string; kg: number; bags: number; count: number; priceTotal: number }
    > = {};
    summaryOffers.forEach((offer) => {
      try {
        const isoDate = new Date(offer.createdAtIso);
        if (Number.isNaN(isoDate.getTime())) return;
        const month = isoDate.getMonth();
        const year = isoDate.getFullYear();
        const key = `${year}-${String(month + 1).padStart(2, "0")}`;
        const label = `${["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"][month]} / ${year}`;
        if (!acc[key]) acc[key] = { key, label, kg: 0, bags: 0, count: 0, priceTotal: 0 };
        acc[key].kg += offer.quantityKg;
        acc[key].bags += offer.quantityKg / 60;
        acc[key].count += 1;
        acc[key].priceTotal += offer.ourPrice;
      } catch {
        // skip
      }
    });
    return Object.values(acc).sort((a, b) => a.key.localeCompare(b.key));
  }, [summaryOffers]);

  const summaryByOrigin = useMemo(() => {
    const acc: Record<string, { name: string; kg: number }> = {};
    summaryOffers.forEach((offer) => {
      if (!acc[offer.origin]) acc[offer.origin] = { name: offer.origin, kg: 0 };
      acc[offer.origin].kg += offer.quantityKg;
    });
    return Object.values(acc).sort((a, b) => b.kg - a.kg);
  }, [summaryOffers]);

  const summaryByCommercial = useMemo(() => {
    const acc: Record<string, { name: string; kg: number }> = {};
    summaryOffers.forEach((offer) => {
      if (!acc[offer.commercial]) acc[offer.commercial] = { name: offer.commercial, kg: 0 };
      acc[offer.commercial].kg += offer.quantityKg;
    });
    return Object.values(acc).sort((a, b) => b.kg - a.kg);
  }, [summaryOffers]);

  const summaryByStatus = useMemo(() => {
    const acc: Record<string, { name: string; count: number; kg: number }> = {};
    offers.forEach((offer) => {
      if (!acc[offer.status]) acc[offer.status] = { name: offer.status, count: 0, kg: 0 };
      acc[offer.status].count += 1;
      acc[offer.status].kg += offer.quantityKg;
    });
    return Object.values(acc).sort((a, b) => b.count - a.count);
  }, [offers]);

  // ---------- Render branches ----------
  if (!authReady) {
    return (
      <main className="min-h-screen bg-slate-100 px-4 py-10">
        <div className="mx-auto max-w-md rounded-2xl bg-white p-8 text-center shadow-lg">
          <p className="text-slate-600">Carregando sistema...</p>
        </div>
      </main>
    );
  }

  if (!userEmail) {
    return (
      <main className="min-h-screen bg-slate-100 px-4 py-10">
        <div className="mx-auto max-w-md rounded-2xl bg-white p-8 shadow-lg">
          <p className="text-sm font-semibold uppercase tracking-wider text-emerald-700">
            Book de Ofertas
          </p>

          <h1 className="mt-2 text-3xl font-bold text-slate-900">
            {authMode === "signin"
              ? "Entrar no sistema"
              : authMode === "signup"
              ? "Criar conta"
              : "Recuperar senha"}
          </h1>

          <p className="mt-3 text-sm text-slate-600">
            {authMode === "reset"
              ? "Informe seu e-mail para receber o link de redefinição."
              : "Use seu e-mail e senha cadastrados no sistema."}
          </p>

          <form className="mt-8 space-y-5" onSubmit={handleAuthSubmit}>
            <TextField
              label="E-mail"
              value={authEmail}
              onChange={setAuthEmail}
              placeholder="seuemail@empresa.com"
              type="email"
              required
            />

            {authMode !== "reset" && (
              <TextField
                label="Senha"
                value={authPassword}
                onChange={setAuthPassword}
                placeholder="Sua senha"
                type="password"
                required
                min="6"
              />
            )}

            <button
              type="submit"
              disabled={authBusy}
              className="w-full rounded-lg bg-emerald-700 px-4 py-3 font-semibold text-white hover:bg-emerald-800 disabled:cursor-not-allowed disabled:opacity-60"
            >
              {authBusy
                ? "Enviando..."
                : authMode === "signin"
                ? "Entrar"
                : authMode === "signup"
                ? "Criar conta"
                : "Enviar link de recuperação"}
            </button>
          </form>

          {authMessage && (
            <p className="mt-4 rounded-lg bg-emerald-50 p-3 text-sm text-emerald-800">
              {authMessage}
            </p>
          )}

          {authError && (
            <p className="mt-4 rounded-lg bg-red-50 p-3 text-sm text-red-800">
              {authError}
            </p>
          )}

          <div className="mt-6 flex flex-col gap-2 text-sm text-slate-600">
            {authMode !== "signin" && (
              <button
                type="button"
                onClick={() => {
                  setAuthMode("signin");
                  setAuthError("");
                  setAuthMessage("");
                }}
                className="text-left text-emerald-700 hover:underline"
              >
                Já tenho conta — entrar
              </button>
            )}
            {authMode !== "signup" && (
              <button
                type="button"
                onClick={() => {
                  setAuthMode("signup");
                  setAuthError("");
                  setAuthMessage("");
                }}
                className="text-left text-emerald-700 hover:underline"
              >
                Quero criar uma conta
              </button>
            )}
            {authMode !== "reset" && (
              <button
                type="button"
                onClick={() => {
                  setAuthMode("reset");
                  setAuthError("");
                  setAuthMessage("");
                }}
                className="text-left text-emerald-700 hover:underline"
              >
                Esqueci minha senha
              </button>
            )}
          </div>
        </div>
      </main>
    );
  }

  if (!dataReady) {
    return (
      <main className="min-h-screen bg-slate-100 px-4 py-10">
        <div className="mx-auto max-w-md rounded-2xl bg-white p-8 text-center shadow-lg">
          <p className="text-slate-600">Carregando ofertas do servidor...</p>
          {dataError && (
            <p className="mt-4 rounded-lg bg-red-50 p-3 text-sm text-red-800">
              {dataError}
            </p>
          )}
        </div>
      </main>
    );
  }

  if (dataError && screen === "dashboard") {
    return (
      <main className="min-h-screen bg-slate-100 px-4 py-10">
        <div className="mx-auto max-w-md rounded-2xl bg-white p-8 shadow-lg">
          <h1 className="text-2xl font-bold text-red-700">
            Não foi possível carregar os dados
          </h1>
          <p className="mt-3 text-sm text-slate-700">{dataError}</p>
          <button
            type="button"
            onClick={() => window.location.reload()}
            className="mt-5 rounded-lg bg-emerald-700 px-4 py-3 font-semibold text-white"
          >
            Tentar de novo
          </button>
        </div>
      </main>
    );
  }

  if (screen === "form") {
    const editing = editingId !== "";

    return (
      <main className="min-h-screen bg-slate-100 px-4 py-8">
        <div className="mx-auto max-w-5xl">
          <div className="flex items-center justify-between">
            <h1 className="text-3xl font-bold text-slate-900">
              {editing ? "Editar oferta" : "Nova oferta"}
            </h1>

            <button
              type="button"
              onClick={() => setScreen("dashboard")}
              className="rounded-lg border border-slate-300 bg-white px-4 py-2"
            >
              Voltar
            </button>
          </div>

          <form
            onSubmit={saveOffer}
            className="mt-6 rounded-2xl bg-white p-6 shadow-sm"
          >
            <div className="grid gap-5 md:grid-cols-2">
              <SelectField
                label="Produto"
                value={product}
                onChange={setProduct}
                options={["Soja", "Milho"]}
              />

              <TextField label="Safra" value={crop} onChange={setCrop} required />

              <SelectField
                label="Originação"
                value={origin}
                onChange={setOrigin}
                options={[
                  "Norte MT",
                  "Médio Norte MT",
                  "Sul MT",
                  "Oeste MT",
                  "Goiás",
                  "Paraná",
                  "Pará",
                  "Tocantins",
                ]}
              />

              <TextField
                label="Fornecedor / ofertante"
                value={supplier}
                onChange={setSupplier}
                required
              />

              <TextField
                label="Local de embarque"
                value={shippingLocation}
                onChange={setShippingLocation}
                required
              />

              <TextField
                label="Destino"
                value={destination}
                onChange={setDestination}
                required
              />

              <SelectField
                label="Comercial responsável"
                value={commercial}
                onChange={setCommercial}
                options={["Antonio", "André", "Nelson", "Clovis"]}
              />

              <div>
                <TextField
                  label="Quantidade em kg"
                  value={quantityKg}
                  onChange={setQuantityKg}
                  type="number"
                  min="1"
                  required
                />
                <p className="mt-1 text-xs text-slate-500">
                  Equivale a {numberFormat(bags)} sacas de 60 kg.
                </p>
              </div>

              <TextField
                label="Preço pedido por saca"
                value={askingPrice}
                onChange={setAskingPrice}
                type="number"
                min="0"
                step="0.01"
                required
              />

              <TextField
                label="Nosso preço por saca"
                value={ourPrice}
                onChange={setOurPrice}
                type="number"
                min="0"
                step="0.01"
              />

              <SelectField
                label="Retirada"
                value={withdrawalType}
                onChange={setWithdrawalType}
                options={["FOB", "CIF"]}
              />

              <SelectField
                label="Qualidade"
                value={quality}
                onChange={setQuality}
                options={["Padrão", "Avariado"]}
              />

              {quality === "Avariado" && (
                <TextField
                  label="Porcentagem de avariado"
                  value={damagedPercentage}
                  onChange={setDamagedPercentage}
                  type="number"
                  min="0"
                  max="100"
                  step="0.01"
                  required
                  placeholder="Exemplo: 5"
                />
              )}

              <SelectField
                label="Status"
                value={status}
                onChange={setStatus}
                options={[
                  "Em negociação",
                  "Enviada",
                  "Fechada",
                  "Perdida",
                  "Cancelada",
                ]}
              />
            </div>

            <div className="mt-5">
              <label className="mb-2 block text-sm font-medium text-slate-700">
                Observações
              </label>

              <textarea
                value={notes}
                onChange={(event) => setNotes(event.target.value)}
                rows={4}
                className="w-full rounded-lg border border-slate-300 px-4 py-3"
              />
            </div>

            <button className="mt-6 rounded-lg bg-emerald-700 px-6 py-3 font-semibold text-white hover:bg-emerald-800">
              {editing ? "Salvar alterações" : "Salvar oferta"}
            </button>
          </form>
        </div>
      </main>
    );
  }

  if (screen === "details" && selectedOffer) {
    const canEdit = selectedOffer.createdById === userId;

    return (
      <main className="min-h-screen bg-slate-100 px-4 py-8">
        <div className="mx-auto max-w-5xl">
          <div className="flex flex-wrap items-center justify-between gap-3">
            <div>
              <p className="text-sm font-semibold text-emerald-700">
                Oferta {selectedOffer.code}
              </p>
              <h1 className="text-3xl font-bold text-slate-900">
                Detalhes da oferta
              </h1>
            </div>

            <div className="flex gap-3">
              {canEdit && (
                <button
                  type="button"
                  onClick={() => editOffer(selectedOffer)}
                  className="rounded-lg bg-emerald-700 px-4 py-2 font-semibold text-white"
                >
                  Editar
                </button>
              )}

              {selectedOffer.status === "Fechada" && (
                <button
                  type="button"
                  onClick={() => openBoleta(selectedOffer)}
                  className="rounded-lg bg-blue-700 px-4 py-2 font-semibold text-white"
                >
                  Gerar boleta
                </button>
              )}

              <button
                type="button"
                onClick={() => setScreen("dashboard")}
                className="rounded-lg border border-slate-300 bg-white px-4 py-2"
              >
                Voltar
              </button>
            </div>
          </div>

          <section className="mt-6 rounded-2xl bg-white p-6 shadow-sm">
            <div className="grid gap-5 md:grid-cols-3">
              <p><strong>Produto:</strong> {selectedOffer.product}</p>
              <p><strong>Safra:</strong> {selectedOffer.crop}</p>
              <p><strong>Status:</strong> {selectedOffer.status}</p>
              <p><strong>Fornecedor:</strong> {selectedOffer.supplier}</p>
              <p><strong>Originação:</strong> {selectedOffer.origin}</p>
              <p><strong>Destino:</strong> {selectedOffer.destination}</p>
              <p><strong>Quantidade:</strong> {numberFormat(selectedOffer.quantityKg)} kg</p>
              <p><strong>Preço pedido:</strong> {money(selectedOffer.askingPrice)}</p>
              <p><strong>Nosso preço:</strong> {money(selectedOffer.ourPrice)}</p>
              <p><strong>Criada por:</strong> {selectedOffer.createdById === userId ? userEmail : "Membro da equipe"}</p>
              <p><strong>Data e hora:</strong> {selectedOffer.createdAt}</p>
              <p><strong>Comercial:</strong> {selectedOffer.commercial}</p>
            </div>
          </section>

          <section className="mt-6 rounded-2xl bg-white p-6 shadow-sm">
            <h2 className="text-xl font-bold text-slate-900">
              Histórico da oferta
            </h2>

            <div className="mt-4 space-y-3">
              {[...selectedOffer.history].reverse().map((item, index) => (
                <div
                  key={`${item.dateTime}-${index}`}
                  className="rounded-lg bg-slate-50 p-4"
                >
                  <p className="font-semibold text-slate-900">{item.action}</p>
                  <p className="text-sm text-slate-600">
                    {item.person} — {item.dateTime}
                  </p>
                </div>
              ))}
            </div>
          </section>
        </div>
      </main>
    );
  }

  if (screen === "summary") {
    const totalBags = summaryVolume / 60;
    const averagePrice = summaryOffers.length
      ? summaryOffers.reduce((total, offer) => total + (offer.ourPrice || 0), 0) /
        summaryOffers.length
      : 0;
    return (
      <main className="min-h-screen bg-slate-100 px-4 py-8">
        <div className="mx-auto max-w-6xl">
          <div className="flex flex-wrap items-center justify-between gap-3">
            <div>
              <p className="text-sm font-semibold text-emerald-700">
                Book de Ofertas
              </p>
              <h1 className="text-3xl font-bold text-slate-900">
                Resumo geral das negociações
              </h1>
              <p className="mt-1 text-sm text-slate-600">
                Indicadores calculados com todas as ofertas cadastradas no
                sistema.
              </p>
            </div>
            <button
              type="button"
              onClick={() => setScreen("dashboard")}
              className="rounded-lg border border-slate-300 bg-white px-4 py-2"
            >
              Voltar ao painel
            </button>
          </div>

          <section className="mt-6 grid gap-4 sm:grid-cols-2 lg:grid-cols-4">
            <div className="rounded-xl bg-white p-5 shadow-sm">
              <p className="text-sm text-slate-500">Volume total negociado</p>
              <p className="mt-2 text-2xl font-bold text-slate-900">
                {numberFormat(summaryVolume)} kg
              </p>
            </div>
            <div className="rounded-xl bg-white p-5 shadow-sm">
              <p className="text-sm text-slate-500">Total de sacas</p>
              <p className="mt-2 text-2xl font-bold text-slate-900">
                {numberFormat(totalBags)}
              </p>
            </div>
            <div className="rounded-xl bg-white p-5 shadow-sm">
              <p className="text-sm text-slate-500">Total de negociações</p>
              <p className="mt-2 text-2xl font-bold text-slate-900">
                {summaryOffers.length}
              </p>
            </div>
            <div className="rounded-xl bg-white p-5 shadow-sm">
              <p className="text-sm text-slate-500">Preço médio nosso</p>
              <p className="mt-2 text-2xl font-bold text-slate-900">
                {money(averagePrice)}
              </p>
            </div>
          </section>

          <section className="mt-6 overflow-hidden rounded-2xl bg-white shadow-sm">
            <h2 className="bg-blue-900 px-6 py-4 text-lg font-bold text-white">
              Resumo por mês e ano
            </h2>
            <div className="overflow-x-auto">
              <table className="w-full min-w-190 text-left text-sm">
                <thead className="bg-slate-50 text-xs uppercase text-slate-500">
                  <tr>
                    <th className="px-6 py-4">Mês / ano</th>
                    <th className="px-6 py-4">Total kg</th>
                    <th className="px-6 py-4">Total sacas</th>
                    <th className="px-6 py-4">Negociações</th>
                    <th className="px-6 py-4">Preço médio nosso</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {summaryByMonth.map((item) => (
                    <tr key={item.label}>
                      <td className="px-6 py-4 font-semibold">{item.label}</td>
                      <td className="px-6 py-4">{numberFormat(item.kg)} kg</td>
                      <td className="px-6 py-4">{numberFormat(item.bags)}</td>
                      <td className="px-6 py-4">{item.count}</td>
                      <td className="px-6 py-4">
                        {money(item.priceTotal / item.count)}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </section>

          <div className="mt-6 grid gap-6 lg:grid-cols-2">
            <section className="overflow-hidden rounded-2xl bg-white shadow-sm">
              <h2 className="bg-blue-900 px-6 py-4 text-lg font-bold text-white">
                Volume por originação
              </h2>
              <table className="w-full text-left text-sm">
                <thead className="bg-slate-50 text-xs uppercase text-slate-500">
                  <tr>
                    <th className="px-6 py-4">Originação</th>
                    <th className="px-6 py-4">Total kg</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {summaryByOrigin.map((item) => (
                    <tr key={item.name}>
                      <td className="px-6 py-4">{item.name}</td>
                      <td className="px-6 py-4 font-semibold">
                        {numberFormat(item.kg)} kg
                      </td>
                    </tr>
                  ))}
                  <tr className="bg-emerald-50 font-bold text-emerald-900">
                    <td className="px-6 py-4">
                      TOTAL DE TODAS AS ORIGINAÇÕES
                    </td>
                    <td className="px-6 py-4">
                      {numberFormat(summaryVolume)} kg
                    </td>
                  </tr>
                </tbody>
              </table>
            </section>
            <section className="overflow-hidden rounded-2xl bg-white shadow-sm">
              <h2 className="bg-blue-900 px-6 py-4 text-lg font-bold text-white">
                Volume por comercial
              </h2>
              <table className="w-full text-left text-sm">
                <thead className="bg-slate-50 text-xs uppercase text-slate-500">
                  <tr>
                    <th className="px-6 py-4">Comercial</th>
                    <th className="px-6 py-4">Total kg</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {summaryByCommercial.map((item) => (
                    <tr key={item.name}>
                      <td className="px-6 py-4">{item.name}</td>
                      <td className="px-6 py-4 font-semibold">
                        {numberFormat(item.kg)} kg
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </section>
          </div>

          <section className="mt-6 overflow-hidden rounded-2xl bg-white shadow-sm">
            <h2 className="bg-blue-900 px-6 py-4 text-lg font-bold text-white">
              Resumo de status das negociações
            </h2>
            <div className="overflow-x-auto">
              <table className="w-full min-w-155 text-left text-sm">
                <thead className="bg-slate-50 text-xs uppercase text-slate-500">
                  <tr>
                    <th className="px-6 py-4">Status</th>
                    <th className="px-6 py-4">Quantidade de negociações</th>
                    <th className="px-6 py-4">Volume total (kg)</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {summaryByStatus.map((item) => (
                    <tr key={item.name}>
                      <td className="px-6 py-4">
                        <span
                          className={`rounded-full px-3 py-1 text-xs font-semibold ${statusColor(item.name)}`}
                        >
                          {item.name}
                        </span>
                      </td>
                      <td className="px-6 py-4">{item.count}</td>
                      <td className="px-6 py-4 font-semibold">
                        {numberFormat(item.kg)} kg
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </section>
        </div>
      </main>
    );
  }

  if (screen === "boleta" && selectedOffer) {
    const bagsInOffer = selectedOffer.quantityKg / 60;
    const grossValue = bagsInOffer * selectedOffer.ourPrice;
    const commissionTotal =
      bagsInOffer *
      (Number(boleta.commissionOneValue || 0) +
        Number(boleta.commissionTwoValue || 0));
    const titleClass =
      "mt-8 rounded-t-lg bg-blue-900 px-5 py-3 text-lg font-bold text-white";
    const sectionClass =
      "rounded-b-lg border border-slate-200 bg-white p-5 shadow-sm";

    return (
      <main className="min-h-screen bg-slate-100 px-4 py-8">
        <div className="mx-auto max-w-5xl">
          <div className="flex flex-wrap items-center justify-between gap-3">
            <div>
              <p className="text-sm font-semibold text-emerald-700">
                Oferta fechada: {selectedOffer.code}
              </p>
              <h1 className="text-3xl font-bold text-slate-900">
                Preencher boleta de compra / venda
              </h1>
              <p className="mt-1 text-sm text-slate-600">
                Preencha os campos abaixo na mesma ordem em que aparecerão no
                PDF.
              </p>
            </div>
            <button
              type="button"
              onClick={() => setScreen("details")}
              className="rounded-lg border border-slate-300 bg-white px-4 py-2"
            >
              Voltar à oferta
            </button>
          </div>

          <div className="mt-6">
            <h2 className={titleClass}>1. IDENTIFICAÇÃO DA OPERAÇÃO</h2>
            <section className={sectionClass}>
              <div className="grid gap-5 md:grid-cols-2">
                <TextField label="Nº da boleta" value={selectedOffer.code} onChange={() => {}} />
                <TextField label="Data / hora do fechamento" value={selectedOffer.createdAt} onChange={() => {}} />
                <SelectField
                  label="Tipo de negociação"
                  value={boleta.negotiationType}
                  onChange={(v) => updateBoleta("negotiationType", v)}
                  options={[
                    "Compra no MI",
                    "Compra no ME",
                    "Venda no MI",
                    "Venda no ME",
                    "Troca / Barter",
                    "Fixação de preço",
                    "Washout",
                  ]}
                />
                <TextField label="Comercial responsável" value={selectedOffer.commercial} onChange={() => {}} />
              </div>
            </section>

            <h2 className={titleClass}>2. CONTRAPARTE</h2>
            <section className={sectionClass}>
              <div className="grid gap-5 md:grid-cols-2">
                <TextField label="Comprador" value={boleta.buyer} onChange={(v) => updateBoleta("buyer", v)} required />
                <TextField label="CNPJ do comprador" value={boleta.buyerDocument} onChange={(v) => updateBoleta("buyerDocument", v)} placeholder="00.000.000/0000-00" />
                <TextField label="Vendedor (razão social)" value={selectedOffer.supplier} onChange={() => {}} />
                <TextField label="CNPJ / CPF do vendedor" value={boleta.sellerDocument} onChange={(v) => updateBoleta("sellerDocument", v)} />
                <TextField label="Inscrição estadual" value={boleta.sellerStateRegistration} onChange={(v) => updateBoleta("sellerStateRegistration", v)} />
                <TextField label="Município / UF" value={boleta.sellerCityState} onChange={(v) => updateBoleta("sellerCityState", v)} placeholder="Ex.: Sorriso / MT" />
                <TextField label="Contato responsável" value={boleta.contactPerson} onChange={(v) => updateBoleta("contactPerson", v)} />
                <TextField label="Telefone" value={boleta.contactPhone} onChange={(v) => updateBoleta("contactPhone", v)} />
                <TextField label="E-mail para assinatura do contrato" value={boleta.contractEmail} onChange={(v) => updateBoleta("contractEmail", v)} type="email" />
              </div>
            </section>

            <h2 className={titleClass}>3. PRODUTO E QUALIDADE</h2>
            <section className={sectionClass}>
              <div className="grid gap-5 md:grid-cols-2">
                <TextField label="Produto" value={selectedOffer.product} onChange={() => {}} />
                <TextField label="Safra" value={selectedOffer.crop} onChange={() => {}} />
                <TextField
                  label="Classificação / padrão"
                  value={
                    selectedOffer.quality === "Avariado"
                      ? `Avariado (${selectedOffer.damagedPercentage}%)`
                      : selectedOffer.quality
                  }
                  onChange={() => {}}
                />
              </div>
            </section>

            <h2 className={titleClass}>4. QUANTIDADE E PREÇO</h2>
            <section className={sectionClass}>
              <div className="grid gap-5 md:grid-cols-2">
                <TextField label="Quantidade contratada (sacas 60 kg)" value={numberFormat(bagsInOffer)} onChange={() => {}} />
                <TextField label="Quantidade (kg)" value={numberFormat(selectedOffer.quantityKg)} onChange={() => {}} />
                <TextField label="Preço unitário / saca (R$)" value={money(selectedOffer.ourPrice)} onChange={() => {}} />
                <TextField label="VALOR BRUTO DA MERCADORIA (R$)" value={money(grossValue)} onChange={() => {}} />
              </div>
            </section>

            <h2 className={titleClass}>5. FRETE, COMISSÕES E TRIBUTOS</h2>
            <section className={sectionClass}>
              <div className="grid gap-5 md:grid-cols-2">
                <TextField label="Modalidade do frete" value={selectedOffer.withdrawalType} onChange={() => {}} />
                <SelectField label="Imposto por conta de" value={boleta.taxResponsible} onChange={(v) => updateBoleta("taxResponsible", v)} options={["Vendedor", "Comprador"]} />
                <SelectField label="Operação com PIS / COFINS" value={boleta.pisCofins} onChange={(v) => updateBoleta("pisCofins", v)} options={["N/A", "Sim", "Não"]} />
                <SelectField label="Funrural — folha ou produção" value={boleta.funrural} onChange={(v) => updateBoleta("funrural", v)} options={["N/A", "Folha de pagamento", "Produção"]} />
                <TextField label="Outras retenções (FETHAB/FUNRURAL/IAGRO)" value={boleta.otherWithholdings} onChange={(v) => updateBoleta("otherWithholdings", v)} />
                <TextField label="Base de cálculo da comissão" value="Por saca (R$/sc)" onChange={() => {}} />
                <TextField label="Comissão 1 — beneficiário" value={boleta.commissionOneName} onChange={(v) => updateBoleta("commissionOneName", v)} />
                <TextField label="Comissão 1 — CPF / CNPJ" value={boleta.commissionOneDocument} onChange={(v) => updateBoleta("commissionOneDocument", v)} />
                <TextField label="Comissão 1 — valor R$ por saca" value={boleta.commissionOneValue} onChange={(v) => updateBoleta("commissionOneValue", v)} type="number" min="0" step="0.01" />
                <TextField label="Comissão 2 — beneficiário" value={boleta.commissionTwoName} onChange={(v) => updateBoleta("commissionTwoName", v)} />
                <TextField label="Comissão 2 — CPF / CNPJ" value={boleta.commissionTwoDocument} onChange={(v) => updateBoleta("commissionTwoDocument", v)} />
                <TextField label="Comissão 2 — valor R$ por saca" value={boleta.commissionTwoValue} onChange={(v) => updateBoleta("commissionTwoValue", v)} type="number" min="0" step="0.01" />
                <TextField label="TOTAL DE COMISSÕES (R$)" value={money(commissionTotal)} onChange={() => {}} />
              </div>
            </section>

            <h2 className={titleClass}>6. PAGAMENTO</h2>
            <section className={sectionClass}>
              <div className="grid gap-5 md:grid-cols-2">
                <TextField label="Condição de pagamento" value={boleta.paymentTerms} onChange={(v) => updateBoleta("paymentTerms", v)} />
                <TextField label="Data prevista de pagamento" value={boleta.paymentDate} onChange={(v) => updateBoleta("paymentDate", v)} type="date" />
                <SelectField label="Forma de pagamento" value={boleta.paymentMethod} onChange={(v) => updateBoleta("paymentMethod", v)} options={["PIX", "TED", "Boleto", "Compensação / encontro de contas", "Barter"]} />
                <TextField label="Prazo para emissão da NF" value={boleta.invoiceDeadline} onChange={(v) => updateBoleta("invoiceDeadline", v)} />
                <TextField label="Banco" value={boleta.bank} onChange={(v) => updateBoleta("bank", v)} />
                <TextField label="Agência" value={boleta.bankBranch} onChange={(v) => updateBoleta("bankBranch", v)} />
                <TextField label="Conta corrente" value={boleta.bankAccount} onChange={(v) => updateBoleta("bankAccount", v)} />
                <TextField label="Chave PIX" value={boleta.pixKey} onChange={(v) => updateBoleta("pixKey", v)} />
                <TextField label="Adiantamento (R$)" value={boleta.advancePayment} onChange={(v) => updateBoleta("advancePayment", v)} type="number" min="0" step="0.01" />
                <TextField label="Data do adiantamento" value={boleta.advancePaymentDate} onChange={(v) => updateBoleta("advancePaymentDate", v)} type="date" />
                <TextField label="Garantia do adiantamento" value={boleta.advanceGuarantee} onChange={(v) => updateBoleta("advanceGuarantee", v)} />
                <SelectField label="Titularidade da conta confere?" value={boleta.accountOwnershipConfirmed} onChange={(v) => updateBoleta("accountOwnershipConfirmed", v)} options={["Sim", "Não", "N/A"]} />
                <SelectField label="Certidão de penhor" value={boleta.pledgeCertificate} onChange={(v) => updateBoleta("pledgeCertificate", v)} options={["Sim", "Não", "N/A"]} />
                <SelectField label="Sintegra consultado" value={boleta.sintegraChecked} onChange={(v) => updateBoleta("sintegraChecked", v)} options={["Sim", "Não", "N/A"]} />
                <SelectField label="CAR válido" value={boleta.carValid} onChange={(v) => updateBoleta("carValid", v)} options={["Sim", "Não", "N/A"]} />
              </div>
            </section>

            <h2 className={titleClass}>7. EMBARQUE E LOGÍSTICA</h2>
            <section className={sectionClass}>
              <div className="grid gap-5 md:grid-cols-2">
                <TextField label="Local de embarque" value={selectedOffer.shippingLocation} onChange={() => {}} />
                <TextField label="Município / UF de embarque" value={boleta.shippingCityState} onChange={(v) => updateBoleta("shippingCityState", v)} placeholder="Ex.: Sorriso / MT" />
                <TextField label="Data início do embarque" value={boleta.shippingStart} onChange={(v) => updateBoleta("shippingStart", v)} type="date" />
                <TextField label="Data fim do embarque" value={boleta.shippingEnd} onChange={(v) => updateBoleta("shippingEnd", v)} type="date" />
                <TextField label="Cadência (ton/dia)" value={boleta.dailyTons} onChange={(v) => updateBoleta("dailyTons", v)} type="number" min="0" />
                <TextField
                  label="Dias estimados de embarque"
                  value={
                    boleta.dailyTons
                      ? String(Math.ceil(selectedOffer.quantityKg / 1000 / Number(boleta.dailyTons)))
                      : ""
                  }
                  onChange={() => {}}
                />
              </div>
            </section>

            <div className="mt-8 rounded-lg bg-emerald-50 p-5">
              <p className="font-semibold text-emerald-900">
                Boleta pronta para PDF
              </p>
              <p className="mt-1 text-sm text-emerald-800">
                Salve os dados para reutilizá-los nesta boleta; o PDF também
                salva automaticamente.
              </p>
              <div className="mt-4 flex flex-wrap gap-3">
                <button
                  type="button"
                  onClick={persistBoleta}
                  className="rounded-lg border border-emerald-700 px-6 py-3 font-semibold text-emerald-800 hover:bg-emerald-100"
                >
                  Salvar dados da boleta
                </button>
                <button
                  type="button"
                  onClick={downloadPdf}
                  className="rounded-lg bg-emerald-700 px-6 py-3 font-semibold text-white hover:bg-emerald-800"
                >
                  Baixar boleta em PDF
                </button>
              </div>
              {dataError && (
                <p className="mt-3 rounded-lg bg-red-50 p-3 text-sm text-red-800">
                  {dataError}
                </p>
              )}
            </div>
          </div>
        </div>
      </main>
    );
  }

  return (
    <main className="min-h-screen bg-slate-100 px-4 py-8">
      <div className="mx-auto max-w-6xl">
        <header className="rounded-2xl bg-emerald-800 p-7 text-white shadow-lg">
          <div className="flex flex-col gap-5 lg:flex-row lg:items-center lg:justify-between">
            <div>
              <p className="text-sm font-semibold uppercase tracking-wider text-emerald-200">
                Book de Ofertas
              </p>

              <h1 className="mt-1 text-3xl font-bold">Painel de ofertas</h1>

              <p className="mt-2 text-emerald-100">
                Usuário: {userEmail}
              </p>
            </div>

            <div className="flex flex-wrap gap-3">
              <button
                type="button"
                onClick={() => setScreen("summary")}
                className="rounded-lg border border-emerald-300 px-5 py-3 font-semibold text-white"
              >
                Resumo geral
              </button>
              <button
                type="button"
                onClick={exportBackup}
                className="rounded-lg border border-emerald-300 px-5 py-3 font-semibold text-white"
              >
                Exportar backup
              </button>
              <label className="cursor-pointer rounded-lg border border-emerald-300 px-5 py-3 font-semibold text-white">
                Importar backup
                <input
                  type="file"
                  accept=".xlsx,.xls"
                  onChange={(e) => {
                    importBackupToLocal();
                    e.target.value = "";
                  }}
                  className="hidden"
                />
              </label>

              <button
                type="button"
                onClick={openNewOffer}
                className="rounded-lg bg-white px-5 py-3 font-semibold text-emerald-800"
              >
                + Nova oferta
              </button>

              <button
                type="button"
                onClick={handleSignOut}
                className="rounded-lg border border-emerald-300 px-5 py-3 font-semibold text-white"
              >
                Sair
              </button>
            </div>
          </div>
        </header>

        <section className="mt-6 grid gap-4 md:grid-cols-2">
          <div className="rounded-xl bg-white p-5 shadow-sm">
            <p className="text-sm font-medium text-slate-500">
              Total de ofertas
            </p>
            <p className="mt-2 text-3xl font-bold text-slate-900">
              {offers.length}
            </p>
          </div>

          <div className="rounded-xl bg-white p-5 shadow-sm">
            <p className="text-sm font-medium text-slate-500">
              Volume ofertado
            </p>
            <p className="mt-2 text-3xl font-bold text-slate-900">
              {numberFormat(totalVolume)} kg
            </p>
          </div>
        </section>

        <section className="mt-6 overflow-hidden rounded-2xl bg-white shadow-sm">
          <div className="border-b border-slate-200 p-6">
            <h2 className="text-xl font-bold text-slate-900">
              Ofertas recentes
            </h2>

            <label className="mt-4 block text-sm font-medium text-slate-700">
              Pesquisar oferta pelo código
            </label>

            <input
              type="search"
              value={search}
              onChange={(event) => setSearch(event.target.value)}
              placeholder="Exemplo: OF-001"
              className="mt-2 w-full max-w-md rounded-lg border border-slate-300 px-4 py-3"
            />
          </div>

          <div className="overflow-x-auto">
            <table className="w-full min-w-275 text-left text-sm">
              <thead className="bg-slate-50 text-xs uppercase text-slate-500">
                <tr>
                  <th className="px-5 py-4">Código</th>
                  <th className="px-5 py-4">Produto</th>
                  <th className="px-5 py-4">Fornecedor</th>
                  <th className="px-5 py-4">Criada por</th>
                  <th className="px-5 py-4">Comercial</th>
                  <th className="px-5 py-4">Data de criação</th>
                  <th className="px-5 py-4">Status</th>
                  <th className="px-5 py-4 text-center">Ver</th>
                </tr>
              </thead>

              <tbody className="divide-y divide-slate-100">
                {visibleOffers.map((offer) => (
                  <tr key={offer.id}>
                    <td className="px-5 py-4 font-semibold text-slate-900">
                      {offer.code}
                    </td>
                    <td className="px-5 py-4">{offer.product}</td>
                    <td className="px-5 py-4">{offer.supplier}</td>
                    <td className="px-5 py-4">{offer.createdById === userId ? userEmail : "Membro da equipe"}</td>
                    <td className="px-5 py-4">{offer.commercial}</td>
                    <td className="px-5 py-4">{offer.createdAt}</td>
                    <td className="px-5 py-4">
                      <span
                        className={`rounded-full px-3 py-1 text-xs font-semibold ${statusColor(offer.status)}`}
                      >
                        {offer.status}
                      </span>
                    </td>
                    <td className="px-5 py-4 text-center">
                      <button
                        type="button"
                        onClick={() => openOffer(offer)}
                        className="rounded-lg border border-emerald-200 px-3 py-2 text-lg text-emerald-700 hover:bg-emerald-50"
                        title="Ver oferta completa"
                      >
                        🔍
                      </button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </section>
      </div>
    </main>
  );
}
