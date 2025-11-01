'use client';

import { motion } from 'framer-motion';
import Link from 'next/link';
import { ReactNode } from 'react';

interface CardProps {
  /** Zawartość karty */
  children: ReactNode;
  /** Opcjonalny link */
  href?: string;
  /** Klasy CSS */
  className?: string;
  /** Indeks animacji */
  index?: number;
}

/**
 * Komponent karty z animacjami i hover effects
 */
export default function Card({ children, href, className = '', index = 0 }: CardProps) {
  const cardContent = (
    <motion.div
      initial={{ opacity: 0, y: 20 }}
      whileInView={{ opacity: 1, y: 0 }}
      whileHover={{ y: -8, scale: 1.02 }}
      viewport={{ once: true, margin: "-50px" }}
      transition={{
        duration: 0.5,
        delay: index * 0.1,
      }}
      className={`
        relative overflow-hidden
        bg-white rounded-xl border border-gray-200
        p-6 cursor-pointer group
        shadow-md hover:shadow-xl
        transition-all duration-300 ease-out
        hover:border-blue-300
        ${className}
      `}
    >
      {/* Simple gradient line */}
      <div className="absolute top-0 left-0 w-full h-1 bg-gradient-to-r from-blue-500 to-blue-600 transform scale-x-0 group-hover:scale-x-100 transition-transform duration-300" />
      
      {/* Content */}
      <div className="relative z-10">
        {children}
      </div>
    </motion.div>
  );

  if (href) {
    return (
      <Link href={href} className="block">
        {cardContent}
      </Link>
    );
  }

  return cardContent;
}