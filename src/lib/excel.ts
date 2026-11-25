import * as XLSX from 'xlsx';
import { supabase } from '@/lib/supabase';
import { accountDisplayLabel } from '@/types';
import type {
  Account,
  Category,
  Card,
  PaymentMethod,
  Transaction,
  Invoice,
  Profile,
  SpaceMember,
  TransactionType,
  TransactionStatus,
  SplitType,
  PaymentMethodKind,
  CategoryKind,
} from '@/types';
import { logger } from '@/lib/logger';
import { lastDayOfMonth } from '@/lib/dates';

interface ExportData {
  accounts: Account[];
  categories: Category[];
  cards: Card[];
  paymentMethods: PaymentMethod[];
  transactions: Transaction[];
  invoices: Invoice[];
  members: SpaceMember[];
}

/* ------------------------------------------------------------------ */
/* Helpers                                                            */
/* ------------------------------------------------------------------ */

function centsToCurrency(cents: number): number {
  return cents / 100;
}

function currencyToCents(value: number | string): number {
  if (value === '' || value === null || value === undefined) return 0;
  const num =
    typeof value === 'string' ? parseFloat(value.replace(/\./g, '').replace(',', '.')) : value;
  if (isNaN(num)) return 0;
  return Math.round(num * 100);
}

function formatDateISO(dateStr: string | null | undefined): string {
  if (!dateStr) return '';
  return dateStr.split('T')[0];
}

/** Parse a date that may be YYYY-MM-DD, DD/MM/YYYY, MM/YYYY, or YYYY-MM. */
function parseFlexibleDate(raw: string): string | null {
  if (!raw) return null;
  const s = String(raw).trim();
  // DD/MM/YYYY
  const ddmmyyyy = s.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (ddmmyyyy) return `${ddmmyyyy[3]}-${ddmmyyyy[2]}-${ddmmyyyy[1]}`;
  // YYYY-MM-DD
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  // YYYY-MM -> period only
  if (/^\d{4}-\d{2}$/.test(s)) return null;
  // MM/YYYY
  const mmyyyy = s.match(/^(\d{2})\/(\d{4})$/);
  if (mmyyyy) return null;
  return null;
}

/** Returns {year, month} (month is 1-12) from "YYYY-MM" or "MM/YYYY" or a full date. */
function parsePeriod(raw: string): { year: number; month: number } | null {
  if (!raw) return null;
  const s = String(raw).trim();
  let y: string | null = null;
  let m: string | null = null;
  const ym = s.match(/^(\d{4})-(\d{2})$/);
  if (ym) {
    y = ym[1];
    m = ym[2];
  }
  const my = s.match(/^(\d{2})\/(\d{4})$/);
  if (my) {
    m = my[1];
    y = my[2];
  }
  const full = s.match(/^(\d{4})-(\d{2})-\d{2}$/);
  if (full) {
    y = full[1];
    m = full[2];
  }
  const dmf = s.match(/^\d{2}\/(\d{2})\/(\d{4})$/);
  if (dmf) {
    m = dmf[1];
    y = dmf[2];
  }
  if (y && m) return { year: parseInt(y, 10), month: parseInt(m, 10) };
  return null;
}

function safeSetDate(year: number, monthIndex: number, day: number): Date {
  const last = lastDayOfMonth(year, monthIndex);
  return new Date(year, monthIndex, Math.min(day, last));
}

/** Compute invoice period boundaries given a target period (year, month 1-12) and card closing day. */
function computeInvoicePeriod(
  year: number,
  month1: number,
  closingDay: number,
  dueDay: number,
): { periodStart: string; periodEnd: string; closingDate: string; dueDate: string } {
  // Billing cycle: starts on closing_day of the *previous* month, ends on (closing_day - 1) of current.
  // Simpler model (matches InvoicesPage helpers): period starts at closing_day of previous month,
  // ends at closing_day of current month. We align period_start to the closing_day of the previous month.
  const targetMonthIndex = month1 - 1;
  const startDate = safeSetDate(year, targetMonthIndex - 1, closingDay);
  const endDate = safeSetDate(year, targetMonthIndex, closingDay);
  // closing_date = end; due_date = due_day of the following month relative to end
  const due = safeSetDate(endDate.getFullYear(), endDate.getMonth() + 1, dueDay);
  return {
    periodStart: startDate.toISOString().split('T')[0],
    periodEnd: endDate.toISOString().split('T')[0],
    closingDate: endDate.toISOString().split('T')[0],
    dueDate: due.toISOString().split('T')[0],
  };
}

/** Given an occurred_at date and a card's closing_day, compute the invoice period. */
function periodFromOccurrence(occurredAt: string, closingDay: number, dueDay: number): { periodStart: string; periodEnd: string; closingDate: string; dueDate: string; year: number; month: number } {
  const d = new Date(occurredAt + 'T00:00:00');
  // If the purchase happened on/after closing_day, it falls into next month's billing period.
  // The billing "month" label is the month whose period_end lands in it.
  let year = d.getFullYear();
  let monthIndex = d.getMonth();
  if (d.getDate() >= closingDay) {
    monthIndex += 1;
    if (monthIndex > 11) {
      monthIndex = 0;
      year += 1;
    }
  }
  const period = computeInvoicePeriod(year, monthIndex + 1, closingDay, dueDay);
  return { ...period, year, month: monthIndex + 1 };
}

function normalize(s: string): string {
  return s
    .trim()
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '');
}

function getVal(row: Record<string, unknown>, ...keys: string[]): string {
  for (const k of keys) {
    if (row[k] !== undefined && row[k] !== '') return String(row[k]).trim();
  }
  return '';
}

/* ------------------------------------------------------------------ */
/* Fetch                                                              */
/* ------------------------------------------------------------------ */

export async function fetchAllData(spaceId: string): Promise<ExportData> {
  const [accountsRes, categoriesRes, cardsRes, paymentMethodsRes, transactionsRes, invoicesRes, membersRes] =
    await Promise.all([
      supabase.from('accounts').select('*').eq('space_id', spaceId).order('created_at'),
      supabase.from('categories').select('*').eq('space_id', spaceId).order('name'),
      supabase.from('cards').select('*').eq('space_id', spaceId).order('name'),
      supabase.from('payment_methods').select('*').eq('space_id', spaceId).order('name'),
      supabase
        .from('transactions')
        .select(
          '*, categories:categories!transactions_category_id_fkey(name), accounts:accounts!transactions_account_id_fkey(name, institution, holder_name), destination_accounts:accounts!transactions_destination_account_id_fkey(name, institution, holder_name), destination_members:profiles!transactions_destination_member_id_fkey(display_name), payment_methods:payment_methods!transactions_payment_method_id_fkey(name), cards:cards!transactions_card_id_fkey(name), responsible:profiles!transactions_responsible_id_fkey(display_name), payer:profiles!transactions_payer_id_fkey(display_name), secondary_payer:profiles!transactions_secondary_payer_id_fkey(display_name)',
        )
        .eq('space_id', spaceId)
        .order('occurred_at', { ascending: false }),
      supabase.from('invoices').select('*').eq('space_id', spaceId).order('due_date', { ascending: false }),
      supabase
        .from('space_members')
        .select('*, profiles:profiles!space_members_user_id_fkey(display_name)')
        .eq('space_id', spaceId),
    ]);

  return {
    accounts: (accountsRes.data as Account[]) ?? [],
    categories: (categoriesRes.data as Category[]) ?? [],
    cards: (cardsRes.data as Card[]) ?? [],
    paymentMethods: (paymentMethodsRes.data as PaymentMethod[]) ?? [],
    transactions: (transactionsRes.data as Transaction[]) ?? [],
    invoices: (invoicesRes.data as Invoice[]) ?? [],
    members: (membersRes.data as unknown as SpaceMember[]) ?? [],
  };
}

/* ------------------------------------------------------------------ */
/* Export                                                             */
/* ------------------------------------------------------------------ */

export function exportToExcel(data: ExportData): void {
  const wb = XLSX.utils.book_new();

  // Contas
  const accountsRows = data.accounts.map((a) => ({
    ID: a.id,
    Instituicao: a.institution ?? a.name,
    Titular: a.holder_name ?? '',
    Agencia: a.agency ?? '',
    Numero: a.account_number ?? '',
    Tipo: a.type,
    'Saldo Inicial': centsToCurrency(a.initial_balance_cents),
    'Data Referencia': formatDateISO(a.reference_date),
    Status: a.status,
  }));
  const wsAccounts = XLSX.utils.json_to_sheet(
    accountsRows.length > 0
      ? accountsRows
      : [{ ID: '', Instituicao: '', Titular: '', Agencia: '', Numero: '', Tipo: '', 'Saldo Inicial': '', 'Data Referencia': '', Status: '' }],
  );
  XLSX.utils.book_append_sheet(wb, wsAccounts, 'Contas');

  // Categorias
  const categoryNameById = new Map<string, string>(data.categories.map((c) => [c.id, c.name]));
  const categoriesRows = data.categories.map((c) => ({
    ID: c.id,
    Nome: c.name,
    'Categoria Pai': c.parent_id ? categoryNameById.get(c.parent_id) ?? '' : '',
    Tipo: c.kind,
    Icone: c.icon ?? '',
    Cor: c.color ?? '',
    Status: c.status,
  }));
  const wsCategories = XLSX.utils.json_to_sheet(
    categoriesRows.length > 0
      ? categoriesRows
      : [{ ID: '', Nome: '', 'Categoria Pai': '', Tipo: '', Icone: '', Cor: '', Status: '' }],
  );
  XLSX.utils.book_append_sheet(wb, wsCategories, 'Categorias');

  // Formas de Pagamento
  const pmRows = data.paymentMethods.map((pm) => ({
    ID: pm.id,
    Nome: pm.name,
    Tipo: pm.kind,
    Padrao: pm.is_default ? 'Sim' : 'Nao',
    Status: pm.status,
  }));
  const wsPM = XLSX.utils.json_to_sheet(
    pmRows.length > 0 ? pmRows : [{ ID: '', Nome: '', Tipo: '', Padrao: '', Status: '' }],
  );
  XLSX.utils.book_append_sheet(wb, wsPM, 'Formas de Pagamento');

  // Cartoes
  const cardsRows = data.cards.map((c) => ({
    ID: c.id,
    Nome: c.name,
    Bandeira: c.brand ?? '',
    Titular: c.holder_name ?? '',
    'Ultimos Digitos': c.last_digits ?? '',
    Limite: centsToCurrency(c.limit_cents),
    'Dia Fechamento': c.closing_day,
    'Dia Vencimento': c.due_day,
    Status: c.status,
  }));
  const wsCards = XLSX.utils.json_to_sheet(
    cardsRows.length > 0
      ? cardsRows
      : [{ ID: '', Nome: '', Bandeira: '', Titular: '', 'Ultimos Digitos': '', Limite: '', 'Dia Fechamento': '', 'Dia Vencimento': '', Status: '' }],
  );
  XLSX.utils.book_append_sheet(wb, wsCards, 'Cartoes');

  // Lancamentos
  const txRows = data.transactions.map((t) => ({
    ID: t.id,
    Tipo: t.type,
    Descricao: t.description,
    Valor: centsToCurrency(t.amount_cents),
    Categoria: t.categories?.name ?? '',
    'Forma de Pagamento': t.payment_methods?.name ?? '',
    Conta: t.accounts ? accountDisplayLabel(t.accounts as Account) : '',
    'Conta Destino': t.destination_accounts ? accountDisplayLabel(t.destination_accounts as Account) : '',
    'Membro Destino': t.destination_members?.display_name ?? '',
    Cartao: t.cards?.name ?? '',
    Fatura: '',
    'Parcela Atual': '',
    'Total Parcelas': '',
    Estabelecimento: t.establishment ?? '',
    CPF_CNPJ: t.cpf_cnpj ?? '',
    'Data Realizacao': formatDateISO(t.occurred_at),
    Vencimento: formatDateISO(t.due_date),
    'Data Pagamento': formatDateISO(t.paid_at),
    Status: t.status,
    Responsavel: t.responsible?.display_name ?? '',
    Pagador: t.payer?.display_name ?? '',
    'Tipo Divisao': t.split_type,
    'Valor Pagador Principal': t.payer_amount_cents != null ? centsToCurrency(t.payer_amount_cents) : '',
    'Segundo Pagador': t.secondary_payer?.display_name ?? '',
    'Valor Segundo Pagador': t.secondary_payer_amount_cents != null ? centsToCurrency(t.secondary_payer_amount_cents) : '',
    Observacoes: t.notes ?? '',
  }));
  const txCols = LANCAMENTOS_COLUMNS.reduce((acc, c) => ({ ...acc, [c]: '' }), {} as Record<string, string>);
  const wsTx = XLSX.utils.json_to_sheet(txRows.length > 0 ? txRows : [txCols]);
  XLSX.utils.book_append_sheet(wb, wsTx, 'Lancamentos');

  XLSX.writeFile(wb, `dualcash_export_${new Date().toISOString().split('T')[0]}.xlsx`);
}

/* ------------------------------------------------------------------ */
/* Template                                                           */
/* ------------------------------------------------------------------ */

const LANCAMENTOS_COLUMNS = [
  'ID',
  'Tipo',
  'Descricao',
  'Valor',
  'Categoria',
  'Forma de Pagamento',
  'Conta',
  'Conta Destino',
  'Membro Destino',
  'Cartao',
  'Fatura',
  'Parcela Atual',
  'Total Parcelas',
  'Estabelecimento',
  'CPF_CNPJ',
  'Data Realizacao',
  'Vencimento',
  'Data Pagamento',
  'Status',
  'Responsavel',
  'Pagador',
  'Tipo Divisao',
  'Valor Pagador Principal',
  'Segundo Pagador',
  'Valor Segundo Pagador',
  'Observacoes',
];

export function downloadTemplate(): void {
  const wb = XLSX.utils.book_new();

  // Contas
  const accountsCols = [
    { ID: '', Instituicao: 'Banco Exemplo', Titular: 'João', Agencia: '0001', Numero: '12345-6', Tipo: 'checking', 'Saldo Inicial': 1000, 'Data Referencia': new Date().toISOString().split('T')[0], Status: 'active' },
  ];
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(accountsCols), 'Contas');

  // Categorias
  const categoriesCols = [
    { ID: '', Nome: 'Moradia', 'Categoria Pai': '', Tipo: 'expense', Icone: '', Cor: '', Status: 'active' },
    { ID: '', Nome: 'Aluguel', 'Categoria Pai': 'Moradia', Tipo: 'expense', Icone: '', Cor: '', Status: 'active' },
    { ID: '', Nome: 'Salário', 'Categoria Pai': '', Tipo: 'income', Icone: '', Cor: '', Status: 'active' },
  ];
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(categoriesCols), 'Categorias');

  // Formas de Pagamento
  const pmCols = [
    { ID: '', Nome: 'PIX', Tipo: 'expense', Padrao: 'Sim', Status: 'active' },
    { ID: '', Nome: 'Dinheiro', Tipo: 'expense', Padrao: 'Nao', Status: 'active' },
    { ID: '', Nome: 'Cartão de crédito', Tipo: 'expense', Padrao: 'Nao', Status: 'active' },
    { ID: '', Nome: 'Transferência bancária', Tipo: 'expense', Padrao: 'Nao', Status: 'active' },
    { ID: '', Nome: 'Depósito em conta', Tipo: 'income', Padrao: 'Nao', Status: 'active' },
  ];
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(pmCols), 'Formas de Pagamento');

  // Cartoes
  const cardsCols = [
    { ID: '', Nome: 'Visa 1234', Bandeira: 'Visa', Titular: 'João', 'Ultimos Digitos': '1234', Limite: 5000, 'Dia Fechamento': 3, 'Dia Vencimento': 10, Status: 'active' },
  ];
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(cardsCols), 'Cartoes');

  // Lancamentos - example rows + validation reference rows
  const exampleRows: Record<string, string | number>[] = [
    {
      ID: '', Tipo: 'Despesa', Descricao: 'Supermercado', Valor: 150.0, Categoria: 'Moradia',
      'Forma de Pagamento': 'PIX', Conta: 'Banco Exemplo - João', 'Conta Destino': '', 'Membro Destino': '',
      Cartao: '', Fatura: '', 'Parcela Atual': 1, 'Total Parcelas': 1,
      Estabelecimento: 'Mercadinho', CPF_CNPJ: '', 'Data Realizacao': new Date().toISOString().split('T')[0],
      Vencimento: '', 'Data Pagamento': '', Status: 'Pago', Responsavel: 'João', Pagador: 'João',
      'Tipo Divisao': 'Individual', 'Valor Pagador Principal': '', 'Segundo Pagador': '', 'Valor Segundo Pagador': '',
      Observacoes: 'Exemplo de despesa paga no PIX',
    },
    {
      ID: '', Tipo: 'Despesa', Descricao: 'Notebook', Valor: 3000.0, Categoria: 'Moradia',
      'Forma de Pagamento': 'Cartão de crédito', Conta: '', 'Conta Destino': '', 'Membro Destino': '',
      Cartao: 'Visa 1234', Fatura: '', 'Parcela Atual': 1, 'Total Parcelas': 6,
      Estabelecimento: 'Loja Tech', CPF_CNPJ: '12.345.678/0001-90', 'Data Realizacao': new Date().toISOString().split('T')[0],
      Vencimento: '', 'Data Pagamento': '', Status: 'Pendente', Responsavel: 'João', Pagador: 'João',
      'Tipo Divisao': 'Individual', 'Valor Pagador Principal': '', 'Segundo Pagador': '', 'Valor Segundo Pagador': '',
      Observacoes: 'Exemplo de compra parcelada no cartão',
    },
  ];
  const wsTx = XLSX.utils.json_to_sheet(exampleRows, { header: LANCAMENTOS_COLUMNS });
  XLSX.utils.book_append_sheet(wb, wsTx, 'Lancamentos');

  // Instrucoes sheet
  const instructions = [
    { Campo: 'Tipo', Valores: 'Receita, Despesa, Transferencia' },
    { Campo: 'Status', Valores: 'Pendente, Pago, Parcial, Cancelado' },
    { Campo: 'Tipo Divisao', Valores: 'Individual, Igual, Percentual, Personalizado' },
    { Campo: 'Forma de Pagamento', Valores: 'PIX, Dinheiro, Cartão de crédito, Transferência bancária, Depósito em conta' },
    { Campo: 'Fatura', Valores: 'YYYY-MM (ex: 2026-10), MM/YYYY (ex: 10/2026) ou data DD/MM/AAAA. Em branco = calcula pela Data Realizacao e o Dia Fechamento do cartão.' },
    { Campo: 'Conta', Valores: 'Instituição - Titular (igual à aba Contas). Conta de origem na transferência.' },
    { Campo: 'Conta Destino', Valores: 'Apenas para Transferencia entre contas próprias.' },
    { Campo: 'Membro Destino', Valores: 'Apenas para Transferencia para outro membro do espaço.' },
    { Campo: 'Cartao', Valores: 'Obrigatório quando a Forma de Pagamento for Cartão de crédito.' },
    { Campo: 'Parcela Atual / Total Parcelas', Valores: '1 para à vista. Ex parcelado: 1 e 6.' },
    { Campo: 'Data Pagamento', Valores: 'Se Status = Pago e vazio, será preenchido com a Data Realizacao.' },
    { Campo: 'Categoria Pai', Valores: 'Nome da categoria pai (aba Categorias) para vincular subcategorias.' },
  ];
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(instructions), 'Instrucoes');

  XLSX.writeFile(wb, 'dualcash_modelo.xlsx');
}

/* ------------------------------------------------------------------ */
/* Import - preview + persistence                                     */
/* ------------------------------------------------------------------ */

export interface ImportPreviewRow {
  rowIndex: number;
  type: string;
  description: string;
  amount: string;
  category: string;
  paymentMethod: string;
  account: string;
  card: string;
  invoice: string;
  installment: string;
  status: string;
  resolvedCategory: string;
  resolvedAccount: string;
  resolvedCard: string;
  resolvedInvoice: string;
  errors: string[];
}

export interface ImportPreview {
  rows: ImportPreviewRow[];
  totals: { contas: number; categorias: number; formas: number; cartoes: number; lancamentos: number };
  errors: string[];
}

export interface ImportSummary {
  accounts: number;
  categories: number;
  paymentMethods: number;
  cards: number;
  transactions: number;
  invoices: number;
  invoiceItems: number;
  splits: number;
  errors: string[];
}

type Maps = {
  accountByLabel: Map<string, string>;
  accountByNorm: Map<string, string>;
  categoryByNorm: Map<string, string>;
  pmByNorm: Map<string, string>;
  cardByNorm: Map<string, string>;
  memberByNorm: Map<string, string>;
  profileById: Map<string, Profile>;
};

function buildMaps(existing: ExportData, spaceMembers: SpaceMember[]): Maps {
  const accountByLabel = new Map<string, string>();
  const accountByNorm = new Map<string, string>();
  for (const a of existing.accounts) {
    const label = accountDisplayLabel(a);
    accountByLabel.set(label, a.id);
    accountByNorm.set(normalize(label), a.id);
    accountByNorm.set(normalize(a.institution ?? a.name), a.id);
  }

  const categoryByNorm = new Map<string, string>();
  for (const c of existing.categories) {
    categoryByNorm.set(normalize(c.name), c.id);
  }

  const pmByNorm = new Map<string, string>();
  for (const pm of existing.paymentMethods) {
    pmByNorm.set(normalize(pm.name), pm.id);
  }

  const cardByNorm = new Map<string, string>();
  for (const c of existing.cards) {
    cardByNorm.set(normalize(c.name), c.id);
  }

  const memberByNorm = new Map<string, string>();
  const profileById = new Map<string, Profile>();
  for (const m of spaceMembers) {
    if (m.profiles?.display_name) {
      memberByNorm.set(normalize(m.profiles.display_name), m.user_id);
      profileById.set(m.user_id, { id: m.user_id, display_name: m.profiles.display_name, avatar_url: m.profiles.avatar_url ?? null, created_at: '', updated_at: '' });
    }
  }

  return { accountByLabel, accountByNorm, categoryByNorm, pmByNorm, cardByNorm, memberByNorm, profileById };
}

function resolveAccount(label: string, maps: Maps): string | null {
  if (!label) return null;
  const exact = maps.accountByLabel.get(label);
  if (exact) return exact;
  return maps.accountByNorm.get(normalize(label)) ?? null;
}

function resolveCategory(name: string, maps: Maps): string | null {
  if (!name) return null;
  return maps.categoryByNorm.get(normalize(name)) ?? null;
}

function resolvePaymentMethod(name: string, maps: Maps): string | null {
  if (!name) return null;
  return maps.pmByNorm.get(normalize(name)) ?? null;
}

function resolveCard(name: string, maps: Maps): string | null {
  if (!name) return null;
  return maps.cardByNorm.get(normalize(name)) ?? null;
}

function resolveMember(name: string, maps: Maps): string | null {
  if (!name) return null;
  return maps.memberByNorm.get(normalize(name)) ?? null;
}

const TYPE_MAP: Record<string, TransactionType> = {
  receita: 'income',
  despesa: 'expense',
  transferencia: 'transfer',
  'transferência': 'transfer',
};

const STATUS_MAP: Record<string, TransactionStatus> = {
  pendente: 'pending',
  pago: 'paid',
  parcial: 'partial',
  cancelado: 'cancelled',
};

const SPLIT_MAP: Record<string, SplitType> = {
  individual: 'individual',
  igual: 'equal',
  percentual: 'percent',
  personalizado: 'custom',
};

const KIND_MAP: Record<string, CategoryKind | PaymentMethodKind> = {
  expense: 'expense',
  income: 'income',
  despesa: 'expense',
  receita: 'income',
};

function parseType(raw: string): TransactionType {
  return TYPE_MAP[normalize(raw)] ?? 'expense';
}

function parseStatus(raw: string): TransactionStatus {
  return STATUS_MAP[normalize(raw)] ?? 'pending';
}

function parseSplit(raw: string): SplitType {
  return SPLIT_MAP[normalize(raw)] ?? 'individual';
}

/** Parse the workbook into structured entities + a preview (no DB writes). */
export function parseImportFile(
  fileBuffer: ArrayBuffer,
  existing: ExportData,
  spaceMembers: SpaceMember[],
): ImportPreview {
  const wb = XLSX.read(fileBuffer, { type: 'array' });
  function getRows(sheetName: string): Record<string, unknown>[] {
    const ws = wb.Sheets[sheetName];
    if (!ws) return [];
    return XLSX.utils.sheet_to_json(ws, { defval: '' });
  }

  const maps = buildMaps(existing, spaceMembers);
  const errors: string[] = [];
  const totals = { contas: 0, categorias: 0, formas: 0, cartoes: 0, lancamentos: 0 };

  // Count new entities to be inserted (preview only, no writes)
  // Contas
  const newAccountNorms = new Set<string>();
  for (const row of getRows('Contas')) {
    const institution = getVal(row, 'Instituicao', 'Instituição');
    const holder = getVal(row, 'Titular');
    if (!institution) continue;
    const label = holder ? `${institution} - ${holder}` : institution;
    if (!maps.accountByLabel.has(label) && !newAccountNorms.has(normalize(label))) {
      newAccountNorms.add(normalize(label));
      totals.contas++;
    }
  }

  // Categorias (track names for parent resolution in preview)
  const categoryPreviewNames = new Set<string>(existing.categories.map((c) => normalize(c.name)));
  const newCategoryNorms = new Set<string>();
  for (const row of getRows('Categorias')) {
    const name = getVal(row, 'Nome');
    if (!name) continue;
    const norm = normalize(name);
    if (!categoryPreviewNames.has(norm) && !newCategoryNorms.has(norm)) {
      newCategoryNorms.add(norm);
      categoryPreviewNames.add(norm);
      totals.categorias++;
    }
  }

  // Formas de Pagamento
  const pmPreviewNames = new Set<string>(existing.paymentMethods.map((pm) => normalize(pm.name)));
  const newPmNorms = new Set<string>();
  for (const row of getRows('Formas de Pagamento')) {
    const name = getVal(row, 'Nome');
    if (!name) continue;
    const norm = normalize(name);
    if (!pmPreviewNames.has(norm) && !newPmNorms.has(norm)) {
      newPmNorms.add(norm);
      pmPreviewNames.add(norm);
      totals.formas++;
    }
  }

  // Cartoes
  const cardPreviewNames = new Set<string>(existing.cards.map((c) => normalize(c.name)));
  const newCardNorms = new Set<string>();
  for (const row of getRows('Cartoes')) {
    const name = getVal(row, 'Nome');
    if (!name) continue;
    const norm = normalize(name);
    if (!cardPreviewNames.has(norm) && !newCardNorms.has(norm)) {
      newCardNorms.add(norm);
      cardPreviewNames.add(norm);
      totals.cartoes++;
    }
  }

  // Lancamentos preview
  const txRows = getRows('Lancamentos');
  const txIdSet = new Set(existing.transactions.map((t) => t.id));
  const previewRows: ImportPreviewRow[] = [];

  for (let i = 0; i < txRows.length; i++) {
    const row = txRows[i];
    const type = parseType(getVal(row, 'Tipo'));
    const description = getVal(row, 'Descricao', 'Descrição');
    const valueStr = getVal(row, 'Valor');
    const categoryName = getVal(row, 'Categoria');
    const pmName = getVal(row, 'Forma de Pagamento');
    const accountLabel = getVal(row, 'Conta');
    const destAccountLabel = getVal(row, 'Conta Destino');
    const destMemberName = getVal(row, 'Membro Destino');
    const cardName = getVal(row, 'Cartao', 'Cartão');
    const invoiceRaw = getVal(row, 'Fatura');
    const installmentNum = getVal(row, 'Parcela Atual') || '1';
    const installmentTotal = getVal(row, 'Total Parcelas') || '1';
    const status = parseStatus(getVal(row, 'Status'));
    const occurredAt = getVal(row, 'Data Realizacao', 'Data Realização') || new Date().toISOString().split('T')[0];
    const responsibleName = getVal(row, 'Responsavel', 'Responsável');
    const payerName = getVal(row, 'Pagador');
    const txId = getVal(row, 'ID');

    const rowErrors: string[] = [];
    if (!description) rowErrors.push('Descrição obrigatória');
    if (!valueStr) rowErrors.push('Valor obrigatório');
    if (txId && txIdSet.has(txId)) rowErrors.push('ID já existe (será ignorado)');

    const categoryId = resolveCategory(categoryName, maps);
    const paymentMethodId = resolvePaymentMethod(pmName, maps);
    const accountId = resolveAccount(accountLabel, maps);
    const destAccountId = resolveAccount(destAccountLabel, maps);
    const destMemberId = resolveMember(destMemberName, maps);
    const cardId = resolveCard(cardName, maps);
    const responsibleId = resolveMember(responsibleName, maps);
    const payerId = resolveMember(payerName, maps);

    const isCreditCard = normalize(pmName) === normalize('Cartão de crédito') || (!!cardName && type === 'expense');

    if (categoryName && !categoryId && !newCategoryNorms.has(normalize(categoryName))) {
      rowErrors.push(`Categoria "${categoryName}" não encontrada`);
    }
    if (pmName && !paymentMethodId && !newPmNorms.has(normalize(pmName))) {
      rowErrors.push(`Forma de pagamento "${pmName}" não encontrada`);
    }
    if (accountLabel && !accountId && !newAccountNorms.has(normalize(accountLabel))) {
      rowErrors.push(`Conta "${accountLabel}" não encontrada`);
    }
    if (destAccountLabel && !destAccountId && !newAccountNorms.has(normalize(destAccountLabel))) {
      rowErrors.push(`Conta destino "${destAccountLabel}" não encontrada`);
    }
    if (destMemberName && !destMemberId) {
      rowErrors.push(`Membro destino "${destMemberName}" não encontrado`);
    }
    if (isCreditCard) {
      if (!cardId && !newCardNorms.has(normalize(cardName))) {
        rowErrors.push(`Cartão "${cardName}" não encontrado (obrigatório para crédito)`);
      }
    }
    if (responsibleName && !responsibleId) {
      rowErrors.push(`Responsável "${responsibleName}" não encontrado`);
    }
    if (payerName && !payerId) {
      rowErrors.push(`Pagador "${payerName}" não encontrado`);
    }

    // Determine invoice label for preview
    let invoiceLabel = '';
    if (isCreditCard && (cardId || newCardNorms.has(normalize(cardName)))) {
      const period = parsePeriod(invoiceRaw);
      if (period) {
        invoiceLabel = `${period.year}-${String(period.month).padStart(2, '0')}`;
      } else if (invoiceRaw) {
        invoiceLabel = invoiceRaw;
      } else {
        const card = existing.cards.find((c) => c.id === cardId);
        if (card) {
          const p = periodFromOccurrence(occurredAt, card.closing_day, card.due_day);
          invoiceLabel = `${p.year}-${String(p.month).padStart(2, '0')}`;
        }
      }
    }

    totals.lancamentos++;
    previewRows.push({
      rowIndex: i + 2,
      type: getVal(row, 'Tipo'),
      description,
      amount: valueStr,
      category: categoryName,
      paymentMethod: pmName,
      account: accountLabel,
      card: cardName,
      invoice: invoiceLabel,
      installment: `${installmentNum}/${installmentTotal}`,
      status: getVal(row, 'Status'),
      resolvedCategory: categoryId ? 'OK' : categoryName ? 'NÃO' : '—',
      resolvedAccount: accountId ? 'OK' : accountLabel ? 'NÃO' : '—',
      resolvedCard: cardId ? 'OK' : cardName ? 'NÃO' : '—',
      resolvedInvoice: invoiceLabel,
      errors: rowErrors,
    });
  }

  return { rows: previewRows, totals, errors };
}

/* ------------------------------------------------------------------ */
/* Import - persistence                                               */
/* ------------------------------------------------------------------ */

export async function importFromExcel(
  file: File,
  spaceId: string,
  userId: string,
  existingData: ExportData,
): Promise<ImportSummary> {
  const summary: ImportSummary = {
    accounts: 0,
    categories: 0,
    paymentMethods: 0,
    cards: 0,
    transactions: 0,
    invoices: 0,
    invoiceItems: 0,
    splits: 0,
    errors: [],
  };

  const buffer = await file.arrayBuffer();
  const wb = XLSX.read(buffer, { type: 'array' });
  function getRows(sheetName: string): Record<string, unknown>[] {
    const ws = wb.Sheets[sheetName];
    if (!ws) return [];
    return XLSX.utils.sheet_to_json(ws, { defval: '' });
  }

  // members + maps
  const { data: memberRows } = await supabase
    .from('space_members')
    .select('*, profiles:profiles!space_members_user_id_fkey(display_name, avatar_url)')
    .eq('space_id', spaceId);
  const members = (memberRows as unknown as SpaceMember[]) ?? [];
  const maps = buildMaps(existingData, members);

  // --- Contas ---
  for (const row of getRows('Contas')) {
    const institution = getVal(row, 'Instituicao', 'Instituição');
    const holder = getVal(row, 'Titular');
    if (!institution) {
      summary.errors.push('Conta sem instituição (linha ignorada)');
      continue;
    }
    const label = holder ? `${institution} - ${holder}` : institution;
    if (maps.accountByLabel.has(label)) continue;
    if (maps.accountByNorm.has(normalize(label))) continue;

    const { data, error } = await supabase
      .from('accounts')
      .insert({
        space_id: spaceId,
        name: label,
        holder_name: holder || null,
        institution,
        agency: getVal(row, 'Agencia', 'Agência') || null,
        account_number: getVal(row, 'Numero', 'Número') || null,
        type: (getVal(row, 'Tipo') || 'checking') as Account['type'],
        initial_balance_cents: currencyToCents(getVal(row, 'Saldo Inicial') || '0'),
        reference_date: getVal(row, 'Data Referencia', 'Data Referência') || new Date().toISOString().split('T')[0],
        owner_id: userId,
      })
      .select('id')
      .single();
    if (error) {
      summary.errors.push(`Erro ao importar conta ${label}: ${error.message}`);
    } else {
      maps.accountByLabel.set(label, data.id);
      maps.accountByNorm.set(normalize(label), data.id);
      summary.accounts++;
    }
  }

  // --- Categorias (resolve parent_id by name, two-pass) ---
  const categoryRows = getRows('Categorias');
  const nameToInsertedId = new Map<string, string>();
  // first pass: categories without parent
  for (const row of categoryRows) {
    const parentName = getVal(row, 'Categoria Pai');
    if (parentName) continue; // handle in second pass
    const name = getVal(row, 'Nome');
    const kind = (KIND_MAP[normalize(getVal(row, 'Tipo'))] as CategoryKind) ?? 'expense';
    if (!name) {
      summary.errors.push('Categoria sem nome (linha ignorada)');
      continue;
    }
    if (maps.categoryByNorm.has(normalize(name)) || nameToInsertedId.has(normalize(name))) continue;

    const { data, error } = await supabase
      .from('categories')
      .insert({
        space_id: spaceId,
        name,
        kind,
        icon: getVal(row, 'Icone') || null,
        color: getVal(row, 'Cor') || null,
      })
      .select('id')
      .single();
    if (error) {
      summary.errors.push(`Erro ao importar categoria ${name}: ${error.message}`);
    } else {
      maps.categoryByNorm.set(normalize(name), data.id);
      nameToInsertedId.set(normalize(name), data.id);
      summary.categories++;
    }
  }
  // second pass: categories with parent
  for (const row of categoryRows) {
    const parentName = getVal(row, 'Categoria Pai');
    if (!parentName) continue;
    const name = getVal(row, 'Nome');
    const kind = (KIND_MAP[normalize(getVal(row, 'Tipo'))] as CategoryKind) ?? 'expense';
    if (!name) continue;
    if (maps.categoryByNorm.has(normalize(name)) || nameToInsertedId.has(normalize(name))) continue;

    const parentId = maps.categoryByNorm.get(normalize(parentName)) ?? null;
    const { data, error } = await supabase
      .from('categories')
      .insert({
        space_id: spaceId,
        name,
        kind,
        icon: getVal(row, 'Icone') || null,
        color: getVal(row, 'Cor') || null,
        parent_id: parentId,
      })
      .select('id')
      .single();
    if (error) {
      summary.errors.push(`Erro ao importar categoria ${name}: ${error.message}`);
    } else {
      maps.categoryByNorm.set(normalize(name), data.id);
      nameToInsertedId.set(normalize(name), data.id);
      summary.categories++;
    }
  }

  // --- Formas de Pagamento ---
  for (const row of getRows('Formas de Pagamento')) {
    const name = getVal(row, 'Nome');
    const kind = (KIND_MAP[normalize(getVal(row, 'Tipo'))] as PaymentMethodKind) ?? 'expense';
    if (!name) {
      summary.errors.push('Forma de pagamento sem nome (linha ignorada)');
      continue;
    }
    if (maps.pmByNorm.has(normalize(name))) continue;

    const isDefault = getVal(row, 'Padrao', 'Padrão').toLowerCase() === 'sim';
    const { data, error } = await supabase
      .from('payment_methods')
      .insert({ space_id: spaceId, name, kind, is_default: isDefault })
      .select('id')
      .single();
    if (error) {
      summary.errors.push(`Erro ao importar forma de pagamento ${name}: ${error.message}`);
    } else {
      maps.pmByNorm.set(normalize(name), data.id);
      summary.paymentMethods++;
    }
  }

  // --- Cartoes ---
  for (const row of getRows('Cartoes')) {
    const name = getVal(row, 'Nome');
    if (!name) {
      summary.errors.push('Cartão sem nome (linha ignorada)');
      continue;
    }
    if (maps.cardByNorm.has(normalize(name))) continue;

    const holderName = getVal(row, 'Titular');
    let ownerId = userId;
    if (holderName) {
      const matched = maps.memberByNorm.get(normalize(holderName));
      if (matched) ownerId = matched;
    }
    const { data, error } = await supabase
      .from('cards')
      .insert({
        space_id: spaceId,
        name,
        brand: getVal(row, 'Bandeira') || null,
        holder_name: holderName || null,
        last_digits: getVal(row, 'Ultimos Digitos', 'Últimos Digitos') || null,
        limit_cents: currencyToCents(getVal(row, 'Limite') || '0'),
        closing_day: parseInt(getVal(row, 'Dia Fechamento') || '1', 10),
        due_day: parseInt(getVal(row, 'Dia Vencimento') || '10', 10),
        owner_id: ownerId,
      })
      .select('id, closing_day, due_day')
      .single();
    if (error) {
      summary.errors.push(`Erro ao importar cartão ${name}: ${error.message}`);
    } else {
      maps.cardByNorm.set(normalize(name), data.id);
      summary.cards++;
    }
  }

  // Load existing invoices for this space (for invoice lookup/creation)
  const { data: existingInvoices } = await supabase
    .from('invoices')
    .select('id, card_id, period_start')
    .eq('space_id', spaceId);
  const invoiceKeyToId = new Map<string, string>();
  for (const inv of (existingInvoices as Pick<Invoice, 'id' | 'card_id' | 'period_start'>[]) ?? []) {
    invoiceKeyToId.set(`${inv.card_id}:${inv.period_start}`, inv.id);
  }

  // cache cards by id for closing/due day
  const cardById = new Map<string, Card>();
  for (const c of existingData.cards) cardById.set(c.id, c);
  // refresh card cache with newly inserted cards (fetch again)
  const { data: refreshedCards } = await supabase.from('cards').select('*').eq('space_id', spaceId);
  for (const c of (refreshedCards as Card[]) ?? []) cardById.set(c.id, c);

  // --- Lancamentos ---
  const txRows = getRows('Lancamentos');
  const txIdSet = new Set(existingData.transactions.map((t) => t.id));

  for (const row of txRows) {
    const type = parseType(getVal(row, 'Tipo'));
    const description = getVal(row, 'Descricao', 'Descrição');
    const valueStr = getVal(row, 'Valor');
    const txId = getVal(row, 'ID');
    if (!description) {
      summary.errors.push('Lançamento sem descrição (linha ignorada)');
      continue;
    }
    if (!valueStr) {
      summary.errors.push(`Lançamento "${description}" sem valor (linha ignorado)`);
      continue;
    }
    if (txId && txIdSet.has(txId)) continue;

    const amountCents = currencyToCents(valueStr);
    const categoryName = getVal(row, 'Categoria');
    const pmName = getVal(row, 'Forma de Pagamento');
    const accountLabel = getVal(row, 'Conta');
    const destAccountLabel = getVal(row, 'Conta Destino');
    const destMemberName = getVal(row, 'Membro Destino');
    const cardName = getVal(row, 'Cartao', 'Cartão');
    const invoiceRaw = getVal(row, 'Fatura');
    const installmentNum = parseInt(getVal(row, 'Parcela Atual') || '1', 10) || 1;
    const installmentTotal = parseInt(getVal(row, 'Total Parcelas') || '1', 10) || 1;
    const status = parseStatus(getVal(row, 'Status'));
    const occurredAt = getVal(row, 'Data Realizacao', 'Data Realização') || new Date().toISOString().split('T')[0];
    const dueDate = getVal(row, 'Vencimento') || null;
    const paidAtRaw = getVal(row, 'Data Pagamento');
    const responsibleName = getVal(row, 'Responsavel', 'Responsável');
    const payerName = getVal(row, 'Pagador');
    const splitType = parseSplit(getVal(row, 'Tipo Divisao'));
    const payerAmountRaw = getVal(row, 'Valor Pagador Principal');
    const secondaryPayerName = getVal(row, 'Segundo Pagador');
    const secondaryPayerAmountRaw = getVal(row, 'Valor Segundo Pagador');
    const notes = getVal(row, 'Observacoes', 'Observações') || null;

    const categoryId = resolveCategory(categoryName, maps);
    const paymentMethodId = resolvePaymentMethod(pmName, maps);
    const accountId = resolveAccount(accountLabel, maps);
    const destAccountId = resolveAccount(destAccountLabel, maps);
    const destMemberId = resolveMember(destMemberName, maps);
    const cardId = resolveCard(cardName, maps);
    const responsibleId = resolveMember(responsibleName, maps);
    const payerId = resolveMember(payerName, maps);
    const secondaryPayerId = resolveMember(secondaryPayerName, maps);

    const isCreditCard = normalize(pmName) === normalize('Cartão de crédito') || (!!cardId && type === 'expense');

    // paid_at logic
    let paidAt: string | null = parseFlexibleDate(paidAtRaw);
    if (!paidAt && status === 'paid') paidAt = occurredAt;

    const payerAmountCents = payerAmountRaw ? currencyToCents(payerAmountRaw) : null;
    const secondaryPayerAmountCents = secondaryPayerAmountRaw ? currencyToCents(secondaryPayerAmountRaw) : null;

    let invoiceItemId: string | null = null;
    let activeInvoiceId: string | null = null;

    // Credit card expense -> invoice + invoice_item
    if (isCreditCard && cardId) {
      const card = cardById.get(cardId);
      if (card) {
        // Determine period
        let periodYear: number;
        let periodMonth: number;
        const explicit = parsePeriod(invoiceRaw);
        if (explicit) {
          periodYear = explicit.year;
          periodMonth = explicit.month;
        } else {
          const p = periodFromOccurrence(occurredAt, card.closing_day, card.due_day);
          periodYear = p.year;
          periodMonth = p.month;
        }
        const { periodStart, periodEnd, closingDate, dueDate: invDue } = computeInvoicePeriod(
          periodYear,
          periodMonth,
          card.closing_day,
          card.due_day,
        );

        // find or create invoice
        activeInvoiceId = invoiceKeyToId.get(`${cardId}:${periodStart}`) ?? null;
        if (!activeInvoiceId) {
          const { data: invData, error: invError } = await supabase
            .from('invoices')
            .insert({
              card_id: cardId,
              space_id: spaceId,
              period_start: periodStart,
              period_end: periodEnd,
              closing_date: closingDate,
              due_date: invDue,
              payment_method_id: null,
              status: 'open',
            })
            .select('id')
            .single();
          if (invError) {
            summary.errors.push(`Erro ao criar fatura para "${description}": ${invError.message}`);
          } else {
            const newInvId = invData.id;
            activeInvoiceId = newInvId;
            invoiceKeyToId.set(`${cardId}:${periodStart}`, newInvId);
            summary.invoices++;
          }
        }

        if (activeInvoiceId) {
          // Create invoice_item first (transaction references it)
          const { data: itemData, error: itemError } = await supabase
            .from('invoice_items')
            .insert({
              invoice_id: activeInvoiceId,
              transaction_id: null,
              description,
              establishment: getVal(row, 'Estabelecimento') || null,
              amount_cents: amountCents,
              installment_number: installmentNum,
              installment_total: installmentTotal,
              category_id: categoryId,
              occurred_at: occurredAt,
            })
            .select('id')
            .single();
          if (itemError) {
            summary.errors.push(`Erro ao criar item de fatura para "${description}": ${itemError.message}`);
          } else {
            invoiceItemId = itemData.id;
            summary.invoiceItems++;
          }
        }
      }
    }

    const payload = {
      space_id: spaceId,
      type,
      description,
      amount_cents: amountCents,
      category_id: categoryId,
      account_id: isCreditCard ? null : accountId,
      responsible_id: responsibleId,
      payer_id: payerId,
      payer_amount_cents: payerAmountCents,
      secondary_payer_id: secondaryPayerId,
      secondary_payer_amount_cents: secondaryPayerAmountCents,
      destination_account_id: type === 'transfer' ? destAccountId : null,
      destination_member_id: type === 'transfer' ? destMemberId : null,
      establishment: getVal(row, 'Estabelecimento') || null,
      cpf_cnpj: getVal(row, 'CPF_CNPJ') || null,
      payment_method_id: paymentMethodId,
      card_id: isCreditCard ? cardId : null,
      invoice_item_id: invoiceItemId,
      occurred_at: occurredAt,
      due_date: dueDate,
      paid_at: paidAt,
      status,
      notes,
      split_type: splitType,
      created_by: userId,
    };

    const { data: txData, error: txError } = await supabase
      .from('transactions')
      .insert(payload)
      .select('id')
      .single();

    if (txError) {
      summary.errors.push(`Erro ao importar lançamento "${description}": ${txError.message}`);
      continue;
    }
    summary.transactions++;

    // Link invoice_item.transaction_id back to the transaction and recalc invoice total
    if (invoiceItemId && activeInvoiceId) {
      await supabase.from('invoice_items').update({ transaction_id: txData.id }).eq('id', invoiceItemId);
      const { data: allItems } = await supabase
        .from('invoice_items')
        .select('amount_cents')
        .eq('invoice_id', activeInvoiceId);
      const newTotal = (allItems ?? []).reduce((s, it) => s + (it.amount_cents ?? 0), 0);
      await supabase.from('invoices').update({ total_cents: newTotal }).eq('id', activeInvoiceId);
    }

    // Splits (shared space only)
    if (splitType !== 'individual' && members.length > 1) {
      const memberIds = members.map((m) => m.user_id);
      // exclude payer as the "owner" share; splits record the OTHER members' shares
      const otherMemberIds = payerId ? memberIds.filter((id) => id !== payerId) : memberIds;
      if (splitType === 'equal' && otherMemberIds.length > 0) {
        const perPerson = Math.floor(amountCents / (otherMemberIds.length + 1));
        const remainder = amountCents - perPerson * (otherMemberIds.length + 1);
        const inserts = otherMemberIds.map((mid, i) => ({
          transaction_id: txData.id,
          member_id: mid,
          share_cents: perPerson + (i === 0 ? remainder : 0),
          share_percent: Number(((perPerson / amountCents) * 100).toFixed(2)),
        }));
        const { error: splitErr } = await supabase.from('transaction_splits').insert(inserts);
        if (splitErr) summary.errors.push(`Erro ao dividir lançamento "${description}": ${splitErr.message}`);
        else summary.splits += inserts.length;
      } else if (splitType === 'custom' || splitType === 'percent') {
        // Use payer_amount_cents / secondary_payer_amount_cents if provided
        const inserts: { transaction_id: string; member_id: string; share_cents: number; share_percent: number | null }[] = [];
        if (secondaryPayerId && secondaryPayerAmountCents != null) {
          inserts.push({
            transaction_id: txData.id,
            member_id: secondaryPayerId,
            share_cents: secondaryPayerAmountCents,
            share_percent: splitType === 'percent' ? Number(((secondaryPayerAmountCents / amountCents) * 100).toFixed(2)) : null,
          });
        }
        if (inserts.length > 0) {
          const { error: splitErr } = await supabase.from('transaction_splits').insert(inserts);
          if (splitErr) summary.errors.push(`Erro ao dividir lançamento "${description}": ${splitErr.message}`);
          else summary.splits += inserts.length;
        }
      }
    }
  }

  logger.info('Importação concluída', summary);
  return summary;
}
