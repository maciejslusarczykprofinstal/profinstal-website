import AiAssistant from "@/components/ui/AiAssistant";
import Breadcrumbs from "@/components/ui/Breadcrumbs";

export default function AiAssistantPage() {
  return (
    <div className="min-h-screen bg-gray-50 py-12">
      <div className="max-w-4xl mx-auto px-4">
        <Breadcrumbs items={[
          { label: 'AI Ekspert CWU' }
        ]} />
        
        <div className="text-center mb-8">
          <h1 className="text-3xl font-bold text-gray-900 mb-4">
            AI Ekspert CWU
          </h1>
          <p className="text-lg text-gray-600 max-w-2xl mx-auto">
            Skonsultuj się z ekspertem AI w zakresie instalacji CWU, modernizacji 
            systemów grzewczych i doboru urządzeń. Zadaj pytanie, a otrzymasz 
            konkretną techniczną odpowiedź.
          </p>
        </div>

        <div className="bg-white rounded-lg shadow-md p-6 mb-8">
          <h2 className="text-xl font-semibold mb-4">Przykładowe pytania:</h2>
          <div className="grid md:grid-cols-2 gap-4">
            <div className="space-y-2">
              <h3 className="font-medium text-gray-900">Obliczenia i dobór:</h3>
              <ul className="text-sm text-gray-600 space-y-1">
                <li>• Jaki kocioł wybrać dla 50 mieszkań?</li>
                <li>• Jak obliczyć straty cyrkulacji?</li>
                <li>• Czy moc 25 kW wystarczy?</li>
                <li>• Jak poprawić efektywność CWU?</li>
              </ul>
            </div>
            <div className="space-y-2">
              <h3 className="font-medium text-gray-900">Modernizacja:</h3>
              <ul className="text-sm text-gray-600 space-y-1">
                <li>• Pompa ciepła czy kocioł gazowy?</li>
                <li>• Jak zmodernizować stary system?</li>
                <li>• Optymalizacja izolacji przewodów</li>
                <li>• Systemy rekuperacji ciepła</li>
              </ul>
            </div>
          </div>
        </div>

        <div className="bg-yellow-50 border border-yellow-200 rounded-lg p-4">
          <div className="flex">
            <svg className="w-5 h-5 text-yellow-600 mt-0.5 mr-3" fill="none" stroke="currentColor" viewBox="0 0 24 24">
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} 
                    d="M13 16h-1v-4h-1m1-4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z" />
            </svg>
            <div>
              <h3 className="text-sm font-medium text-yellow-800">Informacja</h3>
              <p className="text-sm text-yellow-700 mt-1">
                AI Assistant jest dostępny w prawym dolnym rogu ekranu. 
                Kliknij ikonę czatu, aby rozpocząć rozmowę z ekspertem.
              </p>
            </div>
          </div>
        </div>
      </div>

      {/* AI Assistant dostępny globalnie */}
      <AiAssistant />
    </div>
  );
}