import { ResultsDisplayProps } from '@/lib/types';

export default function ResultsDisplay({ results }: ResultsDisplayProps) {
  return (
    <div className="grid md:grid-cols-2 gap-4">
      <div className="p-6 border rounded-2xl">
        <div className="text-sm text-gray-500">Straty całkowite</div>
        <div className="text-3xl font-semibold">{results.loss_pct.toFixed(1)}%</div>
        <div className="text-sm">
          {results.loss_PLN_per_m3.toFixed(2)} zł/m³ • {results.loss_PLN_total.toFixed(2)} zł / okres
        </div>
      </div>
      <div className="p-6 border rounded-2xl">
        <div className="text-sm text-gray-500">Koszty</div>
        <div className="text-sm">Teoria: {results.cost_th_PLN_per_m3.toFixed(2)} zł/m³</div>
        <div className="text-sm">Rzeczywiste: {results.cost_real_PLN_per_m3.toFixed(2)} zł/m³</div>
      </div>
    </div>
  );
}