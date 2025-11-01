import CwuCalculator from "@/components/forms/CwuCalculator";
import Breadcrumbs from "@/components/ui/Breadcrumbs";

export default function CwuCalculatorPage() {
  return (
    <main className="min-h-screen bg-white py-12">
      <div className="mx-auto max-w-4xl px-6">
        <Breadcrumbs items={[
          { label: 'Kalkulator CWU' }
        ]} />
        
        <h1 className="text-3xl font-semibold text-center mb-8">
          Kalkulator CWU
        </h1>
        <CwuCalculator />
      </div>
    </main>
  );
}