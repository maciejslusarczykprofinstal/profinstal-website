import { AuditFormProps } from '@/lib/types';

export default function AuditForm({ form, period, onChange, setPeriod, onSubmit, isLoading = false }: AuditFormProps) {
  return (
    <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
      <input 
        className="border rounded-lg p-3" 
        type="date" 
        value={period.from} 
        onChange={e => setPeriod({...period, from: e.target.value})} 
        disabled={isLoading}
      />
      <input 
        className="border rounded-lg p-3" 
        type="date" 
        value={period.to} 
        onChange={e => setPeriod({...period, to: e.target.value})} 
        disabled={isLoading}
      />
      <input 
        className="border rounded-lg p-3" 
        placeholder="Cena zł/GJ" 
        name="price_GJ_PLN" 
        onChange={onChange}
        disabled={isLoading}
      />
      <input 
        className="border rounded-lg p-3" 
        placeholder="E_DHW_GJ" 
        name="E_DHW_GJ" 
        onChange={onChange}
        disabled={isLoading}
      />
      <input 
        className="border rounded-lg p-3" 
        placeholder="V_DHW_m3" 
        name="V_DHW_m3" 
        onChange={onChange}
        disabled={isLoading}
      />
      <input 
        className="border rounded-lg p-3" 
        placeholder="T_hot °C" 
        name="T_hot_C" 
        defaultValue={form.T_hot_C} 
        onChange={onChange}
        disabled={isLoading}
      />
      <input 
        className="border rounded-lg p-3" 
        placeholder="T_cold °C" 
        name="T_cold_C" 
        defaultValue={form.T_cold_C} 
        onChange={onChange}
        disabled={isLoading}
      />
      <button 
        onClick={onSubmit} 
        disabled={isLoading}
        className={`px-5 py-3 rounded-lg md:col-span-3 ${
          isLoading 
            ? 'bg-gray-400 text-gray-200 cursor-not-allowed' 
            : 'bg-black text-white hover:bg-gray-800'
        }`}
      >
        {isLoading ? 'OBLICZANIE...' : 'OBLICZ STRATY'}
      </button>
    </div>
  );
}