import { COMPANY_INFO } from "@/lib/config";

export default function ContactSection() {
  return (
    <section id="kontakt" className="mx-auto max-w-6xl px-6 py-16">
      <h2 className="text-2xl font-semibold">Kontakt</h2>
      <p className="text-gray-600 mt-2">{COMPANY_INFO.email}</p>
    </section>
  );
}