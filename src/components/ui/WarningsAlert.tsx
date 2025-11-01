import { WarningsAlertProps } from '@/lib/types';

export default function WarningsAlert({ warnings }: WarningsAlertProps) {
  if (warnings.length === 0) return null;

  return (
    <div className="p-4 rounded-lg bg-yellow-50 border text-sm">
      {warnings.map((w, i) => (
        <div key={i}>• {w}</div>
      ))}
    </div>
  );
}