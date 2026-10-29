import { serve } from "https://deno.land/std@0.168.0/http/server.ts"
import { createClient } from 'https://esm.sh/@supabase/supabase-js@2'

serve(async (req) => {
  const corsHeaders = {
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Headers': 'authorization, x-client-info, apikey, content-type',
  }

  if (req.method === 'OPTIONS') {
    return new Response('ok', { headers: corsHeaders })
  }

  try {
    const authHeader = req.headers.get('Authorization')
    if (!authHeader) throw new Error('Não autorizado: Token ausente.')

    const supabaseClient = createClient(
      Deno.env.get('SUPABASE_URL') ?? '',
      Deno.env.get('SUPABASE_ANON_KEY') ?? '',
      { global: { headers: { Authorization: authHeader } } }
    )

    const { data: { user }, error: userError } = await supabaseClient.auth.getUser()
    if (userError || !user) throw new Error('Não autorizado: Sessão inválida.')

    const apiKey = Deno.env.get('GOOGLE_API_KEY')
    const sheetId = Deno.env.get('SPREADSHEET_ID')

    // --- NOVA CAMADA: LISTA BRANCA DE E-MAILS PERMITIDOS ---
    const allowedEmails = [
      'contabilidade@dawadm.net',
      'fiscal1@dawadm.net',
      
    ]

    if (!allowedEmails.includes(user.email?.toLowerCase() ?? '')) {
      throw new Error('Acesso negado: Seu e-mail não possui permissão para visualizar estes dados.')
    }
    // --------------------------------------------------------

    // Abas requeridas pelo parser da sua aplicação
  const ranges = [
      "'Lançamentos'!A1:Z5000",
      "'Lançamentos Projetados'!A1:Z5000",
      "'Controle_Comissões_v2'!A1:Z5000",
      "'Repasses_Comissões'!A1:Z5000"
    ]

    const rangesQuery = ranges.map(r => `ranges=${encodeURIComponent(r)}`).join('&')
    const googleUrl = `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values:batchGet?${rangesQuery}&key=${apiKey}&valueRenderOption=UNFORMATTED_VALUE&dateTimeRenderOption=FORMATTED_STRING`

    const googleRes = await fetch(googleUrl)
    const data = await googleRes.json()

    if (!googleRes.ok) {
      throw new Error(data.error?.message || 'Erro ao consultar Google Sheets')
    }

    // Mapeia o resultado indexado pelo nome da aba
    const sheetsMap: Record<string, any[][]> = {}
    if (data.valueRanges) {
      data.valueRanges.forEach((vr: any) => {
        const rawRange = vr.range || ''
        const sheetName = rawRange.split('!')[0].replace(/^'/, '').replace(/'$/, '')
        sheetsMap[sheetName] = vr.values || []
      })
    }

    return new Response(JSON.stringify(sheetsMap), {
      headers: { ...corsHeaders, 'Content-Type': 'application/json' },
      status: 200,
    })
  } catch (error: any) {
    return new Response(JSON.stringify({ error: error.message }), {
      headers: { ...corsHeaders, 'Content-Type': 'application/json' },
      status: 400,
    })
  }
})