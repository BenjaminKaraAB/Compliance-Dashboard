/* =====================================================================================================
   compliancebox-data — de gedeelde opslag van het Compliance Dashboard

   Waarom deze worker bestaat
   --------------------------
   Tot nu toe stuurde elke browser bij elke wijziging de VOLLEDIGE dataset naar één regel in een
   SharePoint-lijst, zonder te controleren of iemand anders er intussen ook iets in had gezet. Werkten
   twee collega's tegelijk, dan won wie het laatst opsloeg en was het werk van de ander stil verdwenen.

   Hier schrijft iedereen alleen de records weg die hij zelf heeft aangeraakt, en elk record draagt een
   versienummer. Klopt dat versienummer niet meer, dan wordt de schrijfactie geweigerd en krijgt de
   browser de nieuwere versie terug om zijn wijziging opnieuw toe te passen. Twintig mensen tegelijk is
   daarmee geen probleem meer, zolang ze niet exact hetzelfde record op exact hetzelfde moment wijzigen —
   en dan gaat er nog steeds niets verloren, want de verliezer krijgt een melding in plaats van stilte.

   Instellen
   ---------
   1. Workers & Pages → Create → Worker → naam: compliancebox-data
   2. Settings → Bindings → D1 database → variabelenaam: DB → database: compliancebox
   3. Deze code plakken en deployen.

   De database staat in West-Europa (WEUR) omdat er BSN's in de klantgegevens zitten.
===================================================================================================== */

const TOEGESTAAN_DOMEIN = '@administratiebox.nl';

/* Alleen deze adressen mogen de worker aanroepen. Nieuw dashboard-adres erbij? Hier toevoegen. */
const TOEGESTANE_HERKOMST = [
  'https://drop-fd5271a2-be0.benjamin-092.workers.dev',
  'https://benjaminkaraab.github.io',
];

const SOORTEN = ['klant', 'store', 'meta'];

/* --------------------------------------------------------------------------------------------------
   Toegang: hetzelfde principe als in compliancebox-ai — het Microsoft-token van de ingelogde collega
   wordt bij Graph gecontroleerd en het account moet binnen het kantoordomein vallen. Er zijn dus geen
   aparte wachtwoorden of sleutels die rond kunnen slingeren.
   Het resultaat wordt 5 minuten onthouden, zodat niet elke schrijfactie een extra rondje langs
   Microsoft maakt.
-------------------------------------------------------------------------------------------------- */
const tokenCache = new Map();

async function controleerToegang(request) {
  const auth = request.headers.get('Authorization') || '';
  if (!auth.startsWith('Bearer ')) {
    return { ok: false, reden: 'Geen Microsoft-token meegestuurd — log in het dashboard in met je AdministratieBox-account.' };
  }
  const sleutel = auth.slice(-40);
  const nu = Date.now();
  const bekend = tokenCache.get(sleutel);
  if (bekend && bekend.tot > nu) return bekend.uitkomst;

  try {
    const res = await fetch('https://graph.microsoft.com/v1.0/me?$select=userPrincipalName,mail', {
      headers: { Authorization: auth },
    });
    if (!res.ok) {
      return { ok: false, reden: 'Microsoft wees het token af (status ' + res.status + '). Log opnieuw in bij Instellingen.' };
    }
    const wie = await res.json();
    const adres = String(wie.userPrincipalName || wie.mail || '').toLowerCase();
    const uitkomst = adres.endsWith(TOEGESTAAN_DOMEIN)
      ? { ok: true, wie: adres }
      : { ok: false, reden: 'Account valt buiten het toegestane domein.' };
    if (uitkomst.ok) tokenCache.set(sleutel, { uitkomst, tot: nu + 5 * 60 * 1000 });
    if (tokenCache.size > 200) tokenCache.clear();
    return uitkomst;
  } catch (e) {
    return { ok: false, reden: 'Tokencontrole mislukte: ' + (e.message || e) };
  }
}

function corsKoppen(request) {
  const herkomst = request.headers.get('Origin') || '';
  const toegestaan = TOEGESTANE_HERKOMST.includes(herkomst) ? herkomst : TOEGESTANE_HERKOMST[0];
  return {
    'Access-Control-Allow-Origin': toegestaan,
    'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
    'Access-Control-Allow-Headers': 'Authorization, Content-Type',
    'Access-Control-Max-Age': '86400',
    Vary: 'Origin',
  };
}

function json(request, data, status = 200) {
  return new Response(JSON.stringify(data), {
    status,
    headers: { 'Content-Type': 'application/json; charset=utf-8', ...corsKoppen(request) },
  });
}

function geldigRecord(r) {
  return r
    && SOORTEN.includes(r.soort)
    && typeof r.sleutel === 'string' && r.sleutel.length > 0 && r.sleutel.length <= 120
    && typeof r.data === 'string' && r.data.length <= 900000;
}

export default {
  async fetch(request, env) {
    if (request.method === 'OPTIONS') return new Response(null, { status: 204, headers: corsKoppen(request) });

    const url = new URL(request.url);
    const pad = url.pathname.replace(/\/+$/, '') || '/';

    if (!env.DB) {
      return json(request, { fout: 'Geen database gekoppeld. Zet in Settings → Bindings een D1-binding met variabelenaam DB naar de database "compliancebox".' }, 500);
    }

    const toegang = await controleerToegang(request);
    if (!toegang.ok) return json(request, { fout: toegang.reden }, 401);
    const wie = toegang.wie;

    try {
      /* ---------------------------------------------------------------------------------------------
         GET /status — hoeveel staat erin? Gebruikt door het dashboard om te laten zien of de
         overzetting al gedaan is.
      --------------------------------------------------------------------------------------------- */
      if (pad === '/status' && request.method === 'GET') {
        const rij = await env.DB.prepare(
          "SELECT COUNT(*) AS totaal, SUM(soort='klant') AS klanten, SUM(soort='store') AS dossiers, SUM(soort='meta') AS meta, MAX(gewijzigd) AS laatste FROM records"
        ).first();
        return json(request, {
          totaal: rij.totaal || 0,
          klanten: rij.klanten || 0,
          dossiers: rij.dossiers || 0,
          meta: rij.meta || 0,
          laatste: rij.laatste || null,
          jij: wie,
        });
      }

      /* ---------------------------------------------------------------------------------------------
         GET /haal?sinds=<iso> — alles wat sinds dat moment is gewijzigd. Zonder "sinds" krijg je
         alles; dat is alleen bij het opstarten nodig.
      --------------------------------------------------------------------------------------------- */
      if (pad === '/haal' && request.method === 'GET') {
        const sinds = url.searchParams.get('sinds');
        const nu = new Date().toISOString();
        const query = sinds
          ? env.DB.prepare('SELECT soort, sleutel, data, versie, gewijzigd, door FROM records WHERE gewijzigd > ? ORDER BY gewijzigd').bind(sinds)
          : env.DB.prepare('SELECT soort, sleutel, data, versie, gewijzigd, door FROM records ORDER BY gewijzigd');
        const { results } = await query.all();
        return json(request, { nu, aantal: results.length, records: results });
      }

      /* ---------------------------------------------------------------------------------------------
         POST /schrijf — één of meer records wegschrijven.

         Per record wordt gecontroleerd of de versie die de browser meestuurt nog de versie in de
         database is. Zo niet, dan is iemand anders voor geweest: het record komt in "conflicten"
         terug mét de nieuwere inhoud, zodat de browser zijn wijziging daarop opnieuw kan toepassen.
         Niets wordt stil overschreven.

         versie 0 betekent "dit record bestaat nog niet" — dan wordt hij aangemaakt.
      --------------------------------------------------------------------------------------------- */
      if (pad === '/schrijf' && request.method === 'POST') {
        const body = await request.json().catch(() => null);
        const records = body && Array.isArray(body.records) ? body.records : null;
        if (!records || !records.length) return json(request, { fout: 'Geen records meegestuurd.' }, 400);
        if (records.length > 400) return json(request, { fout: 'Te veel records in één keer (maximaal 400).' }, 400);
        for (const r of records) {
          if (!geldigRecord(r)) return json(request, { fout: 'Ongeldig record: ' + JSON.stringify(r && r.sleutel) }, 400);
        }

        const nu = new Date().toISOString();
        const opgeslagen = [];
        const conflicten = [];

        for (const r of records) {
          const verwacht = Number(r.versie) || 0;

          if (verwacht === 0) {
            const res = await env.DB.prepare(
              'INSERT INTO records (soort, sleutel, data, versie, gewijzigd, door) VALUES (?, ?, ?, 1, ?, ?) ON CONFLICT(soort, sleutel) DO NOTHING'
            ).bind(r.soort, r.sleutel, r.data, nu, wie).run();
            if (res.meta.changes === 1) {
              opgeslagen.push({ soort: r.soort, sleutel: r.sleutel, versie: 1 });
            } else {
              const huidig = await env.DB.prepare('SELECT soort, sleutel, data, versie, gewijzigd, door FROM records WHERE soort=? AND sleutel=?').bind(r.soort, r.sleutel).first();
              conflicten.push(huidig);
            }
            continue;
          }

          const res = await env.DB.prepare(
            'UPDATE records SET data=?, versie=versie+1, gewijzigd=?, door=? WHERE soort=? AND sleutel=? AND versie=?'
          ).bind(r.data, nu, wie, r.soort, r.sleutel, verwacht).run();

          if (res.meta.changes === 1) {
            opgeslagen.push({ soort: r.soort, sleutel: r.sleutel, versie: verwacht + 1 });
          } else {
            const huidig = await env.DB.prepare('SELECT soort, sleutel, data, versie, gewijzigd, door FROM records WHERE soort=? AND sleutel=?').bind(r.soort, r.sleutel).first();
            if (huidig) conflicten.push(huidig);
            else conflicten.push({ soort: r.soort, sleutel: r.sleutel, data: null, versie: 0, gewijzigd: nu, door: null, verwijderd: true });
          }
        }

        if (conflicten.length) {
          await env.DB.prepare('INSERT INTO logboek (ts, soort, sleutel, actie, door) VALUES (?, ?, ?, ?, ?)')
            .bind(nu, 'sync', String(conflicten.length), 'conflict', wie).run();
        }

        return json(request, { nu, opgeslagen, conflicten });
      }

      /* ---------------------------------------------------------------------------------------------
         POST /import — de eenmalige overzetting vanuit SharePoint.

         Mag alleen in een lege database, zodat een tweede klik op de knop nooit het werk van een dag
         kan terugdraaien. Opnieuw beginnen kan met POST /leegmaken?bevestig=ja-alles-wissen.
      --------------------------------------------------------------------------------------------- */
      if (pad === '/import' && request.method === 'POST') {
        const body = await request.json().catch(() => null);
        const records = body && Array.isArray(body.records) ? body.records : null;
        if (!records || !records.length) return json(request, { fout: 'Geen records meegestuurd.' }, 400);
        if (records.length > 400) return json(request, { fout: 'Stuur maximaal 400 records per keer.' }, 400);
        for (const r of records) {
          if (!geldigRecord(r)) return json(request, { fout: 'Ongeldig record: ' + JSON.stringify(r && r.sleutel) }, 400);
        }

        const eersteBlok = body.eersteBlok === true;
        if (eersteBlok) {
          const bestaat = await env.DB.prepare('SELECT COUNT(*) AS n FROM records').first();
          if ((bestaat.n || 0) > 0) {
            return json(request, { fout: 'De database is niet leeg — er staat al ' + bestaat.n + ' record(s) in. Overzetten kan maar één keer.' }, 409);
          }
        }

        const nu = new Date().toISOString();
        const stmt = env.DB.prepare('INSERT INTO records (soort, sleutel, data, versie, gewijzigd, door) VALUES (?, ?, ?, 1, ?, ?) ON CONFLICT(soort, sleutel) DO UPDATE SET data=excluded.data, gewijzigd=excluded.gewijzigd, door=excluded.door');
        await env.DB.batch(records.map(r => stmt.bind(r.soort, r.sleutel, r.data, nu, wie)));

        await env.DB.prepare('INSERT INTO logboek (ts, soort, sleutel, actie, door) VALUES (?, ?, ?, ?, ?)')
          .bind(nu, 'import', String(records.length), 'import', wie).run();

        const totaal = await env.DB.prepare('SELECT COUNT(*) AS n FROM records').first();
        return json(request, { geplaatst: records.length, totaal: totaal.n, nu });
      }

      /* ---------------------------------------------------------------------------------------------
         POST /leegmaken?bevestig=ja-alles-wissen — noodrem, om een mislukte overzetting over te doen.
      --------------------------------------------------------------------------------------------- */
      if (pad === '/leegmaken' && request.method === 'POST') {
        if (url.searchParams.get('bevestig') !== 'ja-alles-wissen') {
          return json(request, { fout: 'Bevestiging ontbreekt.' }, 400);
        }
        const voor = await env.DB.prepare('SELECT COUNT(*) AS n FROM records').first();
        await env.DB.prepare('DELETE FROM records').run();
        await env.DB.prepare('INSERT INTO logboek (ts, soort, sleutel, actie, door) VALUES (?, ?, ?, ?, ?)')
          .bind(new Date().toISOString(), 'beheer', String(voor.n || 0), 'leeggemaakt', wie).run();
        return json(request, { gewist: voor.n || 0 });
      }

      return json(request, { fout: 'Onbekend adres. Beschikbaar: /status, /haal, /schrijf, /import, /leegmaken.' }, 404);
    } catch (e) {
      return json(request, { fout: 'Databasefout: ' + (e.message || String(e)) }, 500);
    }
  },
};
