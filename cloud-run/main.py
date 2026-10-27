"""
main.py — HTTP wrapper around the roster optimiser, for Cloud Run.

The service is a pure function: it receives the four sheet tabs as JSON, solves,
and returns the roster and run log as JSON. It never talks to Google, never
touches a database, and stores nothing — so there are no credentials to leak
and nothing to clean up.

Endpoints
  GET  /            health check (also confirms the solver imports)
  POST /solve       {secret, staff, sites, matrix, demand} -> {ok, master_roster,
                    run_log, stats, warnings}

Auth: a shared secret, set as the ROSTER_SECRET environment variable on the
service and stored in Apps Script's Script Properties. Compared in constant time.
"""
import os
import hashlib
import hmac
import logging
import time

from flask import Flask, request, jsonify

import roster_optimiser_month as optimiser

app = Flask(__name__)
logging.basicConfig(level=logging.INFO)

SECRET = os.environ.get('ROSTER_SECRET', '')


def _solver_fingerprint():
    """Short hash of the solver file, so you can confirm which build is live."""
    try:
        path = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                            'roster_optimiser_month.py')
        with open(path, 'rb') as fh:
            return hashlib.sha256(fh.read()).hexdigest()[:10]
    except Exception:
        return 'unknown'


SOLVER_VERSION = _solver_fingerprint()
REQUIRED_TABS = ('staff', 'sites', 'matrix', 'demand')
MAX_MB = float(os.environ.get('MAX_REQUEST_MB', 40))


@app.get('/')
def health():
    return jsonify({'ok': True, 'service': 'apunga-roster-optimiser',
                    'solver': 'cp-sat', 'secret_configured': bool(SECRET),
                    'solver_version': SOLVER_VERSION,
                    'solve_seconds': os.environ.get('SOLVE_SECONDS', '?'),
                    'solve_workers': os.environ.get('SOLVE_WORKERS', '?')})


@app.post('/solve')
def solve():
    started = time.time()

    size_mb = (request.content_length or 0) / 1e6
    if size_mb > MAX_MB:
        return jsonify({'ok': False,
                        'error': f'Request too large ({size_mb:.1f} MB). '
                                 'Is the TravelMatrix tab much bigger than expected?'}), 413

    body = request.get_json(silent=True)
    if not isinstance(body, dict):
        return jsonify({'ok': False, 'error': 'Body must be JSON.'}), 400

    if SECRET:
        supplied = body.get('secret') or request.headers.get('X-Roster-Secret', '')
        if not hmac.compare_digest(str(supplied), SECRET):
            logging.warning('rejected request: bad secret')
            return jsonify({'ok': False, 'error': 'Unauthorised.'}), 401
    else:
        logging.warning('ROSTER_SECRET is not set — the service is unprotected')

    missing = [t for t in REQUIRED_TABS if not body.get(t)]
    if missing:
        return jsonify({'ok': False,
                        'error': 'Missing or empty tab(s): ' + ', '.join(missing)}), 400

    payload = {t: body[t] for t in REQUIRED_TABS}
    rows = {t: len(payload[t]) for t in REQUIRED_TABS}
    logging.info('solving: %s (%.2f MB)', rows, size_mb)

    try:
        result = optimiser.solve(payload, verbose=False)
    except Exception as exc:                       # belt and braces; solve() already guards
        logging.exception('solve crashed')
        return jsonify({'ok': False,
                        'error': f'The optimiser failed unexpectedly: '
                                 f'{type(exc).__name__}: {exc}'}), 500

    elapsed = round(time.time() - started, 1)
    result.setdefault('stats', {})['seconds'] = elapsed
    console = result.pop('console', '')            # keep the response small
    if console:
        logging.info('optimiser console:\n%s', console[:20000])
    logging.info('done in %ss, ok=%s', elapsed, result.get('ok'))

    # A data problem (ok=False) is still a successful HTTP call — the caller gets
    # a run log explaining what to fix, and Apps Script can display it normally.
    return jsonify(result), 200


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 8080)))