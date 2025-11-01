'use client';

import { motion } from 'framer-motion';
import { SERVICES } from "@/lib/config";
import ServiceCard from '@/components/ui/ServiceCard';

export default function ServicesGrid() {
  return (
    <section id="uslugi" className="mx-auto max-w-7xl px-6 py-16">
      {/* Header */}
      <motion.div 
        className="text-center mb-16"
        initial={{ opacity: 0, y: 20 }}
        whileInView={{ opacity: 1, y: 0 }}
        viewport={{ once: true }}
        transition={{ duration: 0.6 }}
      >
        <h2 className="text-4xl font-bold text-gray-900 mb-4">
          Nasze Usługi
        </h2>
        <p className="text-xl text-gray-600 max-w-3xl mx-auto">
          Kompleksowe rozwiązania dla instalacji CWU i systemów grzewczych
        </p>
      </motion.div>

      {/* Services Grid */}
      <motion.div 
        className="grid gap-8 md:grid-cols-2 lg:grid-cols-3"
        initial="hidden"
        whileInView="visible"
        viewport={{ once: true, margin: "-100px" }}
        variants={{
          hidden: {},
          visible: {
            transition: {
              staggerChildren: 0.1
            }
          }
        }}
      >
        {SERVICES.map((service, index) => (
          <ServiceCard
            key={service.title}
            title={service.title}
            description={service.description}
            href={service.href}
            index={index}
          />
        ))}
      </motion.div>
    </section>
  );
}