import Link from "next/link";
import ActiveLink from "@/components/ui/ActiveLink";
import { COMPANY_INFO, MAIN_NAVIGATION } from "@/lib/config";

export default function Header() {
  return (
    <header className="mx-auto max-w-6xl px-6 py-6 flex items-center justify-between">
      <Link href="/" className="font-bold text-xl hover:text-gray-700 transition-colors">
        {COMPANY_INFO.name}
      </Link>
      <nav className="flex gap-6 text-sm">
        {MAIN_NAVIGATION.map((item) => {
          // Sprawdź czy to hash link (do sekcji na stronie)
          const isHashLink = item.href.startsWith('#');
          
          if (isHashLink) {
            return (
              <a 
                key={item.href} 
                href={item.href} 
                className="hover:underline transition-colors"
              >
                {item.label}
              </a>
            );
          }
          
          // Używaj ActiveLink dla routingu z oznaczeniem aktywnej strony
          return (
            <ActiveLink 
              key={item.href} 
              href={item.href}
            >
              {item.label}
            </ActiveLink>
          );
        })}
      </nav>
    </header>
  );
}