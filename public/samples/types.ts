export interface Account {
  id: string;
  type: "checking" | "savings";
  balance: number;
}
