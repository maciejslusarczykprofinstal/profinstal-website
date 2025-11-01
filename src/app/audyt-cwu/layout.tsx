import { AUDIT_CWU_METADATA } from "@/lib/config/metadata";

export const metadata = AUDIT_CWU_METADATA;

export default function AudytCWULayout({
  children,
}: {
  children: React.ReactNode;
}) {
  return children;
}