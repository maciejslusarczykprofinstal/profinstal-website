import { CwuPowerResult } from '@/lib/types';

interface CwuResultsDisplayProps {
  results: CwuPowerResult;
}

export default function CwuResultsDisplay({ results }: CwuResultsDisplayProps) {
  return (
    <div className="mt-8 p-6 bg-gray-50 rounded-lg">
      <h3 className="text-lg font-semibold mb-4">Wyniki obliczeń</h3>
      
      <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
        <div className="bg-white p-4 rounded-lg border">
          <div className="text-sm text-gray-500">Moc podstawowa</div>
          <div className="text-2xl font-bold text-blue-600">
            {results.mocPodstawowa} kW
          </div>
        </div>

        <div className="bg-white p-4 rounded-lg border">
          <div className="text-sm text-gray-500">Straty cyrkulacji</div>
          <div className="text-xl font-semibold text-orange-600">
            +{results.stratyCyrkulacji} kW
          </div>
          <div className="text-xs text-gray-400">
            ({results.procentStrat}%)
          </div>
        </div>

        <div className="bg-white p-4 rounded-lg border border-green-200">
          <div className="text-sm text-gray-500">Moc zamówiona</div>
          <div className="text-3xl font-bold text-green-600">
            {results.mocZamowiona} kW
          </div>
        </div>
      </div>

      <div className="mt-4 text-sm text-gray-600">
        <p>
          <strong>Uwaga:</strong> Obliczenia oparte na średnim zużyciu 120 l/mieszkanie/dobę 
          z współczynnikiem simultaneity 0.3.
        </p>
      </div>
    </div>
  );
}