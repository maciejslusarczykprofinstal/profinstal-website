export default function Home() {
  return (
    <main className="min-h-screen bg-white">
      <header className="mx-auto max-w-6xl px-6 py-6 flex items-center justify-between">
        <div className="font-bold text-xl">PROF INSTAL</div>
        <nav className="flex gap-6 text-sm">
          <a href="#uslugi" className="hover:underline">Usługi</a>
          <a href="#kontakt" className="hover:underline">Kontakt</a>
          <a href="/audyt-cwu" className="hover:underline">Audyt CWU</a>
        </nav>
      </header>

      <section className="mx-auto max-w-6xl px-6 py-16">
        <h1 className="text-4xl md:text-5xl font-semibold leading-tight">
          HVAC | CWU | Audyty energetyczne dla spółdzielni i wspólnot
        </h1>
        <p className="mt-4 text-lg text-gray-600">
          Obliczenia mocy, analiza strat cyrkulacji, raporty DOCX, doradztwo modernizacyjne.
        </p>
        <div className="mt-8 flex gap-4">
          <a href="/audyt-cwu" className="px-5 py-3 rounded-lg bg-black text-white">Uruchom Audyt CWU</a>
          <a href="#kontakt" className="px-5 py-3 rounded-lg border">Kontakt</a>
        </div>
      </section>

      <section id="uslugi" className="mx-auto max-w-6xl px-6 py-12 grid gap-6 md:grid-cols-3">
        <div className="p-6 rounded-2xl border">
          <h3 className="font-semibold text-lg">Audyt CWU</h3>
          <p className="text-gray-600 mt-2">Straty [%], zł/m i rekomendacje na podstawie liczników.</p>
        </div>
        <div className="p-6 rounded-2xl border">
          <h3 className="font-semibold text-lg">HVAC/Serwis</h3>
          <p className="text-gray-600 mt-2">Klimatyzacja, wentylacja, równoważenie instalacji.</p>
        </div>
        <div className="p-6 rounded-2xl border">
          <h3 className="font-semibold text-lg">Raporty DOCX</h3>
          <p className="text-gray-600 mt-2">Wnioski techniczne i finansowe dla zarządów.</p>
        </div>
      </section>

      <section id="kontakt" className="mx-auto max-w-6xl px-6 py-16">
        <h2 className="text-2xl font-semibold">Kontakt</h2>
        <p className="text-gray-600 mt-2">kontakt@profinstal.info</p>
        <p className="text-xs text-gray-400 mt-6">Build test: {new Date().toISOString()}</p>
      </section>
    </main>
  );
}
