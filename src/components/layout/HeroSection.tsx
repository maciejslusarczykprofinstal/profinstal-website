import Link from "next/link";
import { COMPANY_INFO, CTA_BUTTONS } from "@/lib/config";

export default function HeroSection() {
  return (
    <section className="mx-auto max-w-6xl px-6 py-16">
      <h1 className="text-4xl md:text-5xl font-semibold leading-tight">
        {COMPANY_INFO.description}
      </h1>
      <p className="mt-4 text-lg text-gray-600">
        {COMPANY_INFO.tagline}
      </p>
      <div className="mt-8 flex gap-4">
        {CTA_BUTTONS.map((button) => {
          const isHashLink = button.href.startsWith('#');
          
          const buttonClasses = `px-5 py-3 rounded-lg transition-colors ${
            button.variant === 'primary' 
              ? 'bg-black text-white hover:bg-gray-800' 
              : 'border border-gray-300 hover:border-gray-400'
          }`;

          if (isHashLink) {
            return (
              <a 
                key={button.href}
                href={button.href} 
                className={buttonClasses}
              >
                {button.label}
              </a>
            );
          }

          return (
            <Link 
              key={button.href}
              href={button.href} 
              className={buttonClasses}
            >
              {button.label}
            </Link>
          );
        })}
      </div>
    </section>
  );
}