type TimingKey =
  | '1M-1W' | '1M+1W' | '1M'
  | '3M-2W' | '3M+2W' | '3M'
  | '6M-2W' | '6M+2W' | '6M'
  | '12M-1M' | '12M+1M' | '12M'
  | '18M-1M' | '18M+1M' | '18M'
  | '24M-1M' | '24M+1M' | '24M';

type EvalKey = '評価日：1M' | '評価日：3M' | '評価日：6M' | '評価日：12M' | '評価日：18M' | '評価日：24M';

type RecordRow = {
  自動送信OnOff: boolean;
  被験者ID: string;
  難プラID?: string;
  登録日?: Date | string | null;
} & Partial<Record<TimingKey, Date | string | null>> & Partial<Record<EvalKey, Date | string | null>>;

type Settings = {
  登録機関名?: string;
  登録機関コード?: string;
  送信先アドレス?: string; // comma separated allowed
  CC?: string;
  BCC?: string;
  メール件名?: string;
  "メール本文（HTML）"?: string;
  // Optional mappings parsed from 設定シートの表（複数行）
  recipientsByCode?: Record<string, string>; // CODE -> to addresses (comma/semicolon separated)
  namesByCode?: Record<string, string>; // CODE -> institution name
  ccByCode?: Record<string, string>; // CODE -> CC addresses
  bccByCode?: Record<string, string>; // CODE -> BCC addresses
  subjectByCode?: Record<string, string>; // CODE -> メール件名
  bodyByCode?: Record<string, string>; // CODE -> メール本文（HTML） or カスタムメール本文（HTML）
};

type LogRow = {
  被験者ID: string;
  登録医療機関: string;
  送信タイミング: string;
  送信先: string; // To | CC | BCC joined
  送信日時: string; // yyyy/MM/dd HH:mm:ss
  送信結果: string; // 成功 / 失敗：...
};
