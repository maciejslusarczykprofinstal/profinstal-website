'use client';

import { ChevronRightIcon } from '@heroicons/react/24/outline';
import Card from '@/components/ui/Card';

interface ServiceCardProps {
  /** Tytuł usługi */
  title: string;
  /** Opis usługi */
  description: string;
  /** Opcjonalny link */
  href?: string;
  /** Indeks dla animacji */
  index?: number;
}

/**
 * Karta usługi z profesjonalnym designem
 */
export default function ServiceCard({ title, description, href, index = 0 }: ServiceCardProps) {
  return (
    <Card href={href} index={index} className="h-full">
      <div className="flex flex-col h-full">
        {/* Nagłówek */}
        <div className="flex items-start justify-between mb-4">
          <h3 className="text-xl font-bold text-gray-900 group-hover:text-blue-600 transition-colors duration-200">
            {title}
          </h3>
          {href && (
            <ChevronRightIcon className="w-5 h-5 text-gray-400 group-hover:text-blue-600 group-hover:translate-x-1 transition-all duration-200 flex-shrink-0 ml-2" />
          )}
        </div>

        {/* Opis */}
        <p className="text-gray-600 leading-relaxed mb-6 flex-grow">
          {description}
        </p>

        {/* CTA Footer */}
        {href && (
          <div className="pt-4 border-t border-gray-100 group-hover:border-blue-100 transition-colors duration-200">
            <span className="inline-flex items-center text-sm font-semibold text-blue-600 group-hover:text-blue-700 transition-colors duration-200">
              Przejdź do narzędzia
              <ChevronRightIcon className="w-4 h-4 ml-1 group-hover:translate-x-1 transition-transform duration-200" />
            </span>
          </div>
        )}
      </div>
    </Card>
  );
}