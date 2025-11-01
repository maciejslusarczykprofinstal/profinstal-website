import { NextRequest, NextResponse } from 'next/server';
import { calculatePower } from '@/lib/utils/cwu-calculations';
import type { CwuCalculatorData } from '@/lib/types';
import { Document, Packer, Paragraph, Table, TableCell, TableRow, WidthType, AlignmentType, TextRun } from 'docx';

export async function POST(request: NextRequest) {
  try {
    const data: CwuCalculatorData = await request.json();
    const { searchParams } = new URL(request.url);
    const format = searchParams.get('format') || 'json';
    
    // Walidacja wymaganych pól
    if (!data.liczba_mieszkan || !data.liczba_pionow || !data.temp_zimnej_wody || !data.temp_cwu) {
      return NextResponse.json(
        { error: 'Brakuje wymaganych danych' },
        { status: 400 }
      );
    }

    // Obliczenia mocy CWU
    const result = calculatePower(data);
    
    // Generowanie podsumowania
    const summary = {
      liczba_mieszkan: parseInt(data.liczba_mieszkan),
      liczba_pionow: parseInt(data.liczba_pionow),
      temp_zimnej_wody: parseFloat(data.temp_zimnej_wody),
      temp_cwu: parseFloat(data.temp_cwu),
      procent_strat_cyrkulacji: parseFloat(data.procent_strat_cyrkulacji || '0'),
      calculatedAt: new Date().toISOString(),
      recommendations: generateRecommendations(result)
    };

    if (format === 'docx') {
      // Generowanie pliku DOCX
      const docxBuffer = await generateDocxReport(data, result, summary);
      
      return new NextResponse(new Uint8Array(docxBuffer), {
        headers: {
          'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
          'Content-Disposition': `attachment; filename="audyt-cwu-${new Date().toISOString().split('T')[0]}.docx"`,
        },
      });
    }

    // Domyślnie zwracamy JSON
    return NextResponse.json({
      powerKW: result.mocZamowiona,
      summary: summary,
      details: {
        mocPodstawowa: result.mocPodstawowa,
        mocZamowiona: result.mocZamowiona,
        stratyCyrkulacji: result.stratyCyrkulacji,
        procentStrat: result.procentStrat
      }
    });

  } catch (error) {
    console.error('Błąd w API audit-cwu:', error);
    return NextResponse.json(
      { error: 'Błąd podczas przetwarzania danych' },
      { status: 500 }
    );
  }
}

function generateRecommendations(result: { mocZamowiona: number; procentStrat: number }): string[] {
  const recommendations: string[] = [];
  
  if (result.mocZamowiona < 15) {
    recommendations.push('Zalecamy kocioł gazowy o mocy 15-20 kW');
  } else if (result.mocZamowiona < 25) {
    recommendations.push('Zalecamy kocioł gazowy o mocy 20-25 kW');
  } else if (result.mocZamowiona < 35) {
    recommendations.push('Zalecamy kocioł gazowy o mocy 25-35 kW');
  } else {
    recommendations.push('Zalecamy rozważenie kotła o większej mocy lub systemu kaskadowego');
  }

  if (result.procentStrat > 15) {
    recommendations.push('Wysokie straty cyrkulacji - rozważenie izolacji przewodów');
  } else if (result.procentStrat < 5) {
    recommendations.push('Niskie straty cyrkulacji - dobrze zaprojektowana instalacja');
  }

  return recommendations;
}

async function generateDocxReport(
  data: CwuCalculatorData, 
  result: any, 
  summary: any
): Promise<Buffer> {
  const doc = new Document({
    sections: [
      {
        properties: {},
        children: [
          // Nagłówek raportu
          new Paragraph({
            children: [
              new TextRun({
                text: "RAPORT AUDYTU CWU",
                bold: true,
                size: 32,
              }),
            ],
            alignment: AlignmentType.CENTER,
            spacing: { after: 400 },
          }),

          // Data wygenerowania
          new Paragraph({
            children: [
              new TextRun({
                text: `Data wygenerowania: ${new Date().toLocaleDateString('pl-PL')}`,
                size: 20,
              }),
            ],
            spacing: { after: 400 },
          }),

          // Tabela z danymi wejściowymi
          new Paragraph({
            children: [
              new TextRun({
                text: "DANE WEJŚCIOWE",
                bold: true,
                size: 24,
              }),
            ],
            spacing: { before: 200, after: 200 },
          }),

          new Table({
            width: {
              size: 100,
              type: WidthType.PERCENTAGE,
            },
            rows: [
              new TableRow({
                children: [
                  new TableCell({
                    children: [new Paragraph({ children: [new TextRun({ text: "Parametr", bold: true })] })],
                    width: { size: 50, type: WidthType.PERCENTAGE },
                  }),
                  new TableCell({
                    children: [new Paragraph({ children: [new TextRun({ text: "Wartość", bold: true })] })],
                    width: { size: 50, type: WidthType.PERCENTAGE },
                  }),
                ],
              }),
              new TableRow({
                children: [
                  new TableCell({
                    children: [new Paragraph({ children: [new TextRun({ text: "Liczba mieszkań" })] })],
                  }),
                  new TableCell({
                    children: [new Paragraph({ children: [new TextRun({ text: data.liczba_mieszkan })] })],
                  }),
                ],
              }),
              new TableRow({
                children: [
                  new TableCell({
                    children: [new Paragraph({ children: [new TextRun({ text: "Liczba pionów" })] })],
                  }),
                  new TableCell({
                    children: [new Paragraph({ children: [new TextRun({ text: data.liczba_pionow })] })],
                  }),
                ],
              }),
              new TableRow({
                children: [
                  new TableCell({
                    children: [new Paragraph({ children: [new TextRun({ text: "Temperatura zimnej wody (°C)" })] })],
                  }),
                  new TableCell({
                    children: [new Paragraph({ children: [new TextRun({ text: data.temp_zimnej_wody })] })],
                  }),
                ],
              }),
              new TableRow({
                children: [
                  new TableCell({
                    children: [new Paragraph({ children: [new TextRun({ text: "Temperatura CWU (°C)" })] })],
                  }),
                  new TableCell({
                    children: [new Paragraph({ children: [new TextRun({ text: data.temp_cwu })] })],
                  }),
                ],
              }),
              new TableRow({
                children: [
                  new TableCell({
                    children: [new Paragraph({ children: [new TextRun({ text: "Procent strat cyrkulacji (%)" })] })],
                  }),
                  new TableCell({
                    children: [new Paragraph({ children: [new TextRun({ text: data.procent_strat_cyrkulacji || "0" })] })],
                  }),
                ],
              }),
            ],
          }),

          // Wyniki obliczeń
          new Paragraph({
            children: [
              new TextRun({
                text: "WYNIKI OBLICZEŃ",
                bold: true,
                size: 24,
              }),
            ],
            spacing: { before: 400, after: 200 },
          }),

          new Table({
            width: {
              size: 100,
              type: WidthType.PERCENTAGE,
            },
            rows: [
              new TableRow({
                children: [
                  new TableCell({
                    children: [new Paragraph({ children: [new TextRun({ text: "Parametr", bold: true })] })],
                    width: { size: 50, type: WidthType.PERCENTAGE },
                  }),
                  new TableCell({
                    children: [new Paragraph({ children: [new TextRun({ text: "Wartość", bold: true })] })],
                    width: { size: 50, type: WidthType.PERCENTAGE },
                  }),
                ],
              }),
              new TableRow({
                children: [
                  new TableCell({
                    children: [new Paragraph({ children: [new TextRun({ text: "Moc podstawowa (kW)" })] })],
                  }),
                  new TableCell({
                    children: [new Paragraph({ children: [new TextRun({ text: result.mocPodstawowa.toFixed(2) })] })],
                  }),
                ],
              }),
              new TableRow({
                children: [
                  new TableCell({
                    children: [new Paragraph({ children: [new TextRun({ text: "Moc zamówiona (kW)", bold: true, color: "FF0000" })] })],
                  }),
                  new TableCell({
                    children: [new Paragraph({ children: [new TextRun({ text: result.mocZamowiona.toFixed(2), bold: true, color: "FF0000" })] })],
                  }),
                ],
              }),
              new TableRow({
                children: [
                  new TableCell({
                    children: [new Paragraph({ children: [new TextRun({ text: "Straty cyrkulacji (kW)" })] })],
                  }),
                  new TableCell({
                    children: [new Paragraph({ children: [new TextRun({ text: result.stratyCyrkulacji.toFixed(2) })] })],
                  }),
                ],
              }),
              new TableRow({
                children: [
                  new TableCell({
                    children: [new Paragraph({ children: [new TextRun({ text: "Procent strat (%)" })] })],
                  }),
                  new TableCell({
                    children: [new Paragraph({ children: [new TextRun({ text: result.procentStrat.toFixed(1) })] })],
                  }),
                ],
              }),
            ],
          }),

          // Rekomendacje
          new Paragraph({
            children: [
              new TextRun({
                text: "REKOMENDACJE",
                bold: true,
                size: 24,
              }),
            ],
            spacing: { before: 400, after: 200 },
          }),

          ...summary.recommendations.map((rec: string) => 
            new Paragraph({
              children: [
                new TextRun({
                  text: `• ${rec}`,
                  size: 20,
                }),
              ],
              spacing: { after: 100 },
            })
          ),

          // Stopka
          new Paragraph({
            children: [
              new TextRun({
                text: "Raport wygenerowany automatycznie przez system PROF-INSTAL",
                italics: true,
                size: 18,
                color: "666666",
              }),
            ],
            alignment: AlignmentType.CENTER,
            spacing: { before: 600 },
          }),
        ],
      },
    ],
  });

  const buffer = await Packer.toBuffer(doc);
  return buffer;
}