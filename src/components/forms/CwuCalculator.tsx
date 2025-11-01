"use client";

import { useState, useEffect } from "react";
import type { CwuCalculatorData } from "@/lib/types";
import { calculatePower, type CwuPowerResult } from "@/lib/utils/cwu-calculations";
import CwuResultsDisplay from "@/components/ui/CwuResultsDisplay";
import AiAssistant from "@/components/ui/AiAssistant";

const STORAGE_KEY = 'cwu-calculator-data';

export default function CwuCalculator() {
  const [formData, setFormData] = useState<CwuCalculatorData>({
    liczba_mieszkan: "",
    liczba_pionow: "",
    temp_zimnej_wody: "8",
    temp_cwu: "55",
    procent_strat_cyrkulacji: "",
  });

  const [results, setResults] = useState<CwuPowerResult | null>(null);
  const [useApi, setUseApi] = useState(false);
  const [isLoading, setIsLoading] = useState(false);
  const [isDownloading, setIsDownloading] = useState(false);
  const [isDataLoaded, setIsDataLoaded] = useState(false);

  // Wczytanie danych z localStorage przy pierwszym renderze
  useEffect(() => {
    try {
      const savedData = localStorage.getItem(STORAGE_KEY);
      if (savedData) {
        const parsedData = JSON.parse(savedData) as CwuCalculatorData;
        setFormData(parsedData);
        console.log('Wczytano dane z localStorage:', parsedData);
      }
    } catch (error) {
      console.error('Błąd podczas wczytywania danych z localStorage:', error);
    } finally {
      setIsDataLoaded(true);
    }
  }, []);

  // Automatyczne zapisywanie danych do localStorage przy każdej zmianie
  useEffect(() => {
    // Nie zapisuj podczas pierwszego wczytywania, żeby nie nadpisać domyślnych wartości
    if (!isDataLoaded) return;

    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(formData));
      console.log('Zapisano dane do localStorage:', formData);
    } catch (error) {
      console.error('Błąd podczas zapisywania danych do localStorage:', error);
    }
  }, [formData, isDataLoaded]);

  const handleInputChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const { name, value } = e.target;
    setFormData(prev => ({
      ...prev,
      [name]: value
    }));
  };

  const handleCalculate = async () => {
    setIsLoading(true);
    try {
      if (useApi) {
        // Obliczenia przez API
        const response = await fetch('/api/audit-cwu', {
          method: 'POST',
          headers: {
            'Content-Type': 'application/json',
          },
          body: JSON.stringify(formData)
        });

        if (!response.ok) {
          throw new Error('Błąd podczas komunikacji z API');
        }

        const apiResult = await response.json();
        
        // Konwersja wyniku API do formatu CwuPowerResult
        const calculationResults: CwuPowerResult = {
          mocPodstawowa: apiResult.details.mocPodstawowa,
          mocZamowiona: apiResult.details.mocZamowiona,
          stratyCyrkulacji: apiResult.details.stratyCyrkulacji,
          procentStrat: apiResult.details.procentStrat
        };
        
        setResults(calculationResults);
        console.log("Wyniki z API:", apiResult);
      } else {
        // Obliczenia lokalne
        const calculationResults = calculatePower(formData);
        setResults(calculationResults);
        console.log("Wyniki lokalne:", calculationResults);
      }
    } catch (error) {
      console.error("Błąd podczas obliczeń:", error);
      setResults(null);
    } finally {
      setIsLoading(false);
    }
  };

  const handleDownloadReport = async () => {
    if (!results) return;
    
    setIsDownloading(true);
    try {
      const response = await fetch('/api/audit-cwu?format=docx', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify(formData)
      });

      if (!response.ok) {
        throw new Error('Błąd podczas generowania raportu');
      }

      const blob = await response.blob();
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `audyt-cwu-${new Date().toISOString().split('T')[0]}.docx`;
      document.body.appendChild(a);
      a.click();
      window.URL.revokeObjectURL(url);
      document.body.removeChild(a);
    } catch (error) {
      console.error("Błąd podczas pobierania raportu:", error);
    } finally {
      setIsDownloading(false);
    }
  };

  const isFormValid = () => {
    return formData.liczba_mieszkan && 
           formData.liczba_pionow && 
           formData.temp_zimnej_wody && 
           formData.temp_cwu &&
           formData.procent_strat_cyrkulacji;
  };

  const clearSavedData = () => {
    try {
      localStorage.removeItem(STORAGE_KEY);
      setFormData({
        liczba_mieszkan: "",
        liczba_pionow: "",
        temp_zimnej_wody: "8",
        temp_cwu: "55",
        procent_strat_cyrkulacji: "",
      });
      setResults(null);
      console.log('Wyczyszczono zapisane dane');
    } catch (error) {
      console.error('Błąd podczas czyszczenia danych:', error);
    }
  };

  const exportData = () => {
    try {
      const dataToExport = {
        formData,
        exportDate: new Date().toISOString(),
        version: '1.0'
      };
      
      const blob = new Blob([JSON.stringify(dataToExport, null, 2)], {
        type: 'application/json'
      });
      
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `cwu-calculator-data-${new Date().toISOString().split('T')[0]}.json`;
      document.body.appendChild(a);
      a.click();
      URL.revokeObjectURL(url);
      document.body.removeChild(a);
      
      console.log('Dane wyeksportowane');
    } catch (error) {
      console.error('Błąd podczas eksportu danych:', error);
    }
  };

  const importData = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (event) => {
      try {
        const content = event.target?.result as string;
        const importedData = JSON.parse(content);
        
        if (importedData.formData && typeof importedData.formData === 'object') {
          setFormData(importedData.formData);
          setResults(null); // Resetuj wyniki przy imporcie nowych danych
          console.log('Dane zaimportowane:', importedData.formData);
        } else {
          throw new Error('Nieprawidłowy format pliku');
        }
      } catch (error) {
        console.error('Błąd podczas importu danych:', error);
        alert('Błąd podczas importu danych. Sprawdź format pliku.');
      }
    };
    reader.readAsText(file);
    
    // Reset input value, żeby można było importować ten sam plik ponownie
    e.target.value = '';
  };

  return (
    <div className="max-w-2xl mx-auto p-6">
      <div className="flex items-center justify-between mb-6">
        <h2 className="text-2xl font-semibold">Kalkulator CWU</h2>
        
        {/* Wskaźnik autosave i przycisk czyszczenia */}
        <div className="flex items-center space-x-4">
          <div className="flex items-center text-sm text-gray-500">
            <svg className="w-4 h-4 mr-1 text-green-500" fill="none" stroke="currentColor" viewBox="0 0 24 24">
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} 
                    d="M5 13l4 4L19 7" />
            </svg>
            Dane zapisywane automatycznie
          </div>
          
          <button
            onClick={clearSavedData}
            className="text-sm text-gray-400 hover:text-gray-600 transition-colors"
            title="Wyczyść zapisane dane"
          >
            <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} 
                    d="M19 7l-.867 12.142A2 2 0 0116.138 21H7.862a2 2 0 01-1.995-1.858L5 7m5 4v6m4-6v6m1-10V4a1 1 0 00-1-1h-4a1 1 0 00-1 1v3M4 7h16" />
            </svg>
          </button>
        </div>
      </div>
      
      {/* Komunikat o wczytanych danych */}
      {isDataLoaded && (formData.liczba_mieszkan || formData.liczba_pionow) && (
        <div className="mb-4 p-3 bg-blue-50 border border-blue-200 rounded-lg">
          <div className="flex items-center">
            <svg className="w-4 h-4 text-blue-600 mr-2" fill="none" stroke="currentColor" viewBox="0 0 24 24">
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} 
                    d="M13 16h-1v-4h-1m1-4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z" />
            </svg>
            <span className="text-sm text-blue-800">
              Wczytano poprzednio zapisane dane. Możesz kontynuować obliczenia lub wyczyścić dane powyżej.
            </span>
          </div>
        </div>
      )}
      
      <div className="space-y-4">
        <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
          <div>
            <label className="block text-sm font-medium text-gray-700 mb-2">
              Liczba mieszkań
            </label>
            <input
              type="number"
              name="liczba_mieszkan"
              value={formData.liczba_mieszkan}
              onChange={handleInputChange}
              className="w-full border rounded-lg p-3 focus:ring-2 focus:ring-blue-500 focus:border-blue-500"
              placeholder="np. 50"
            />
          </div>

          <div>
            <label className="block text-sm font-medium text-gray-700 mb-2">
              Liczba pionów
            </label>
            <input
              type="number"
              name="liczba_pionow"
              value={formData.liczba_pionow}
              onChange={handleInputChange}
              className="w-full border rounded-lg p-3 focus:ring-2 focus:ring-blue-500 focus:border-blue-500"
              placeholder="np. 4"
            />
          </div>

          <div>
            <label className="block text-sm font-medium text-gray-700 mb-2">
              Temperatura zimnej wody (°C)
            </label>
            <input
              type="number"
              name="temp_zimnej_wody"
              value={formData.temp_zimnej_wody}
              onChange={handleInputChange}
              className="w-full border rounded-lg p-3 focus:ring-2 focus:ring-blue-500 focus:border-blue-500"
              placeholder="8"
            />
          </div>

          <div>
            <label className="block text-sm font-medium text-gray-700 mb-2">
              Temperatura CWU (°C)
            </label>
            <input
              type="number"
              name="temp_cwu"
              value={formData.temp_cwu}
              onChange={handleInputChange}
              className="w-full border rounded-lg p-3 focus:ring-2 focus:ring-blue-500 focus:border-blue-500"
              placeholder="55"
            />
          </div>

          <div className="md:col-span-2">
            <label className="block text-sm font-medium text-gray-700 mb-2">
              Procent strat na cyrkulacji (%)
            </label>
            <input
              type="number"
              name="procent_strat_cyrkulacji"
              value={formData.procent_strat_cyrkulacji}
              onChange={handleInputChange}
              className="w-full border rounded-lg p-3 focus:ring-2 focus:ring-blue-500 focus:border-blue-500"
              placeholder="np. 15"
              step="0.1"
            />
          </div>
        </div>

        <div className="pt-4 border-t">
          <div className="mb-4">
            <label className="flex items-center space-x-2">
              <input
                type="checkbox"
                checked={useApi}
                onChange={(e) => setUseApi(e.target.checked)}
                className="rounded border-gray-300 text-blue-600 focus:ring-blue-500"
              />
              <span className="text-sm text-gray-700">
                Użyj API do obliczeń (separacja backend/frontend)
              </span>
            </label>
          </div>
          
          <button
            onClick={handleCalculate}
            disabled={!isFormValid() || isLoading}
            className={`w-full py-3 px-6 rounded-lg font-semibold transition-colors ${
              isFormValid() && !isLoading
                ? 'bg-black text-white hover:bg-gray-800'
                : 'bg-gray-300 text-gray-500 cursor-not-allowed'
            }`}
          >
            {isLoading ? 'OBLICZAM...' : 'OBLICZ'}
          </button>
          
          {/* Dodatkowe funkcje eksportu/importu */}
          <div className="mt-4 pt-4 border-t border-gray-200">
            <div className="flex justify-center space-x-4">
              <button
                onClick={exportData}
                className="text-sm text-gray-600 hover:text-gray-800 transition-colors flex items-center"
                title="Eksportuj dane do pliku"
              >
                <svg className="w-4 h-4 mr-1" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} 
                        d="M12 10v6m0 0l-3-3m3 3l3-3m2 8H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" />
                </svg>
                Eksportuj dane
              </button>
              
              <label className="text-sm text-gray-600 hover:text-gray-800 transition-colors flex items-center cursor-pointer">
                <svg className="w-4 h-4 mr-1" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} 
                        d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M9 19l3 3m0 0l3-3m-3 3V10" />
                </svg>
                Importuj dane
                <input
                  type="file"
                  accept=".json"
                  onChange={importData}
                  className="hidden"
                />
              </label>
            </div>
          </div>
        </div>
      </div>

      {results && (
        <div className="space-y-4">
          <CwuResultsDisplay results={results} />
          
          <div className="flex justify-center">
            <button
              onClick={handleDownloadReport}
              disabled={isDownloading}
              className={`px-6 py-3 rounded-lg font-semibold transition-colors ${
                isDownloading
                  ? 'bg-gray-300 text-gray-500 cursor-not-allowed'
                  : 'bg-blue-600 text-white hover:bg-blue-700'
              }`}
            >
              {isDownloading ? 'GENERUJĘ RAPORT...' : '📄 POBIERZ RAPORT DOCX'}
            </button>
          </div>
        </div>
      )}
      
      {/* AI Assistant - dostępny zawsze, ale z kontekstem wyników jeśli istnieją */}
      <AiAssistant 
        calculationResults={results || undefined}
        inputData={formData}
      />
    </div>
  );
}