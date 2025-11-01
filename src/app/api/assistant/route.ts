import { NextRequest, NextResponse } from 'next/server';
import OpenAI from 'openai';

const SYSTEM_PROMPT = `Jesteś ekspertem HVAC/CWU (Ciepła Woda Użytkowa). Odpowiadasz konkretnie i technicznie na pytania dotyczące:
- Obliczeń mocy dla instalacji CWU
- Modernizacji systemów grzewczych
- Doboru kotłów i urządzeń grzewczych
- Optymalizacji strat cyrkulacji
- Przepisów i norm technicznych w Polsce
- Rozwiązań energooszczędnych

Używaj polskiej terminologii technicznej. Odpowiadaj krótko i konkretnie, ale dokładnie. Jeśli otrzymasz dane z obliczeń, analizuj je i dawaj praktyczne rekomendacje.`;

export async function POST(request: NextRequest) {
  try {
    const { message, context } = await request.json();
    
    if (!message || typeof message !== 'string') {
      return NextResponse.json(
        { error: 'Brak treści wiadomości' },
        { status: 400 }
      );
    }

    if (!process.env.OPENAI_API_KEY) {
      return NextResponse.json(
        { error: 'Brak konfiguracji API OpenAI. Skonfiguruj OPENAI_API_KEY w zmiennych środowiskowych.' },
        { status: 500 }
      );
    }

    // Inicjalizacja klienta OpenAI tylko gdy mamy API key
    const openai = new OpenAI({
      apiKey: process.env.OPENAI_API_KEY,
    });

    // Przygotowanie kontekstu z wynikami obliczeń
    let contextMessage = '';
    if (context?.calculationResults && context?.inputData) {
      contextMessage = `
KONTEKST OBLICZEŃ CWU:
Dane wejściowe:
- Liczba mieszkań: ${context.inputData.liczba_mieszkan}
- Liczba pionów: ${context.inputData.liczba_pionow}
- Temperatura zimnej wody: ${context.inputData.temp_zimnej_wody}°C
- Temperatura CWU: ${context.inputData.temp_cwu}°C
- Procent strat cyrkulacji: ${context.inputData.procent_strat_cyrkulacji}%

Wyniki obliczeń:
- Moc podstawowa: ${context.calculationResults.mocPodstawowa.toFixed(2)} kW
- Moc zamówiona: ${context.calculationResults.mocZamowiona.toFixed(2)} kW
- Straty cyrkulacji: ${context.calculationResults.stratyCyrkulacji.toFixed(2)} kW
- Procent strat: ${context.calculationResults.procentStrat.toFixed(1)}%

Pytanie użytkownika: ${message}`;
    } else {
      contextMessage = message;
    }

    const completion = await openai.chat.completions.create({
      model: 'gpt-3.5-turbo',
      messages: [
        {
          role: 'system',
          content: SYSTEM_PROMPT
        },
        {
          role: 'user',
          content: contextMessage
        }
      ],
      max_tokens: 500,
      temperature: 0.1, // Niska temperatura dla bardziej deterministycznych odpowiedzi technicznych
    });

    const aiResponse = completion.choices[0]?.message?.content;
    
    if (!aiResponse) {
      throw new Error('Brak odpowiedzi z API OpenAI');
    }

    return NextResponse.json({
      response: aiResponse,
      timestamp: new Date().toISOString()
    });

  } catch (error) {
    console.error('Błąd w API assistant:', error);
    
    if (error instanceof Error && error.message.includes('API key')) {
      return NextResponse.json(
        { error: 'Błąd autoryzacji OpenAI API' },
        { status: 401 }
      );
    }

    return NextResponse.json(
      { error: 'Błąd podczas komunikacji z AI' },
      { status: 500 }
    );
  }
}