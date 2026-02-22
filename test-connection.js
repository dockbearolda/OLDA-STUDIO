/**
 * test-connection.js — Ping interne OLDA Studio
 * Usage : node test-connection.js [base_url]
 * Ex.   : node test-connection.js http://localhost:3000
 *         node test-connection.js https://oldastudio.up.railway.app
 */

const BASE  = process.argv[2] || process.env.TEST_URL || 'http://localhost:3000';
const TOKEN = process.env.DASHBOARD_TOKEN || 'd1e9c4cd170eb57f8ce6254b5aa70b7f708e465d48daefc7f3b0b7e16bd9f4dc';

async function ping() {
    console.log('━━━ Test de connexion OLDA Studio ━━━');
    console.log('  Cible  :', BASE);
    console.log('  Token  :', TOKEN.slice(0, 8) + '…');
    console.log('');

    /* ── 1. Health-check /api/ping (sans auth) ── */
    try {
        const r = await fetch(BASE + '/api/ping');
        const d = await r.json();
        if (d.ok) {
            console.log('✅ Serveur joignable — mode:', d.mode, '— dasholda:', d.dasholda || 'non défini');
        } else {
            console.log('❌ Erreur de liaison : réponse inattendue', d);
        }
    } catch (e) {
        console.error('❌ Erreur de liaison :', e.message);
        process.exit(1);
    }

    /* ── 2. Envoi commande test via relay ── */
    const testOrder = {
        commande  : 'TEST-' + Date.now(),
        date      : new Date().toLocaleDateString('fr-FR'),
        nom       : 'Client Test',
        telephone : '0600000000',
        collection: 'Test',
        reference : 'REF-TEST',
        taille    : 'M',
        couleurTshirt: 'Blanc',
        famille   : 'textile',
        prix      : { tshirt: '25', personnalisation: '0', total: '25' },
        paiement  : { statut: 'NON' }
    };

    try {
        const r = await fetch(BASE + '/api/webhook/forward-to-dasholda', {
            method : 'POST',
            headers: {
                'Content-Type' : 'application/json',
                'Authorization': 'Bearer ' + TOKEN
            },
            body: JSON.stringify(testOrder)
        });
        const d = await r.json().catch(() => ({}));
        if (r.ok && d.ok) {
            console.log('✅ Connexion établie avec le Dashboard — commande:', testOrder.commande);
        } else {
            console.error('❌ Erreur de liaison : [' + r.status + ']', d.error || JSON.stringify(d));
        }
    } catch (e) {
        console.error('❌ Erreur de liaison :', e.message);
    }

    console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
}

ping();
