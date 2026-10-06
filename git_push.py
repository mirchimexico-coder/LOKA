# -*- coding: utf-8 -*-
"""Commit everything in C:\\LOKA with today's date + details, then push to GitHub.

  py git_push.py            (menu 23 - run by hand)
  auto() is called at the end of every loka.refresh_all() (EOD, fixes, counts...)

- Message: "LOKA 05-Oct-2026 | EOD | day rev $X exp $Y net $Z | bal vs capital $W"
  plus a body listing the changed files.
- Never prompts for a password (GIT_TERMINAL_PROMPT=0) and never raises: if the push
  fails (no internet, rejected), the commit stays local and the NEXT run pushes it.
- Git output is sanitised so a token in the remote URL is never printed.
Console is cp1252: plain ASCII output only.
"""
import os, re, sys, subprocess
from datetime import date

REPO = r'C:\LOKA'
REASONS = {('eod.py', None): 'EOD', ('tools.py', 'recount'): 'cash count', ('tools.py', 'move'): 'day fix',
           ('tools.py', 'tips'): 'propinas', ('tools.py', 'settle'): 'owner ledger settle',
           ('tools.py', 'capex'): 'partner-paid expense', ('tools.py', 'repay'): 'partner repayment',
           ('tools.py', 'restore'): 'restore backup', ('loka.py', 'refresh-all'): 'refresh',
           ('git_push.py', None): 'manual push'}

def _clean(s):
    return re.sub(r'https://[^@\s/]+@', 'https://', s or '').strip()

def _git(*args, timeout=90):
    env = dict(os.environ, GIT_TERMINAL_PROMPT='0', GCM_INTERACTIVE='never')
    p = subprocess.run(['git', '-C', REPO, *args], capture_output=True, text=True,
                       encoding='utf-8', errors='replace', timeout=timeout, env=env)
    return p.returncode, _clean(p.stdout), _clean(p.stderr)

def _reason():
    script = os.path.basename(sys.argv[0]).lower() if sys.argv else ''
    arg = sys.argv[1].lower() if len(sys.argv) > 1 else None
    return REASONS.get((script, arg)) or REASONS.get((script, None)) or f'update ({script or "manual"})'

def _details():
    try:
        sys.path.insert(0, REPO)
        import loka
        c = loka.compute(); pos = c['position']
        d, rev, exp, net = c['last_days'][-1]
        return (f"day {d}: rev ${rev:,.0f} exp ${exp:,.0f} net ${net:,.0f}"
                f" | bal vs capital ${pos['balance_vs_capital']:,.0f}")
    except Exception:
        return ''

def run(reason=None, quiet=False):
    """Stage everything, commit with date + details, push. Returns True if pushed."""
    say = (lambda m: print('  github: ' + m)) if quiet else (lambda m: print('   ' + m))
    try:
        rc, out, err = _git('status', '--porcelain', timeout=30)
        if rc != 0:
            say(f'git status failed - skipped ({err[:120]})'); return False
        changed = [l[3:].strip('"') for l in out.splitlines() if l.strip()]
        if changed:
            main = [f for f in changed if not f.startswith(('Backup/', 'days/'))]
            nb = len(changed) - len(main)
            subject = f"LOKA {date.today():%d-%b-%Y} | {reason or _reason()}"
            det = _details()
            if det: subject += f" | {det}"
            body = 'Changed: ' + (', '.join(main) if main else '(backups/day files only)')
            if nb: body += f'\n+ {nb} backup/day file(s)'
            _git('add', '-A', timeout=60)
            rc, out, err = _git('commit', '-m', subject, '-m', body, timeout=60)
            if rc != 0:
                say(f'commit failed - skipped ({(err or out)[:120]})'); return False
            if not quiet: say(f'committed: {subject}')
        # push even when nothing new to commit: an earlier failed push may be waiting
        rc, out, err = _git('status', '-sb', timeout=30)
        if not changed and 'ahead' not in out:
            say('nothing new to push'); return True
        rc, out, err = _git('push', 'origin', 'HEAD', timeout=120)
        if rc == 0:
            say(f'pushed to GitHub ({len(changed)} file(s) changed)'); return True
        hint = 'pull in VS Code first, then retry menu 23' if 'rejected' in err else 'no internet? retry menu 23 later'
        say(f'push FAILED - commit kept locally, {hint}. ({err.splitlines()[-1][:100] if err else ""})')
        return False
    except subprocess.TimeoutExpired:
        say('git timed out - commit kept locally, retry menu 23 later'); return False
    except Exception as e:
        say(f'skipped ({type(e).__name__}: {str(e)[:100]})'); return False

def auto():
    """Called by loka.refresh_all(). One line of output, never raises."""
    return run(quiet=True)

if __name__ == '__main__':
    print("\n== COMMIT + PUSH TO GITHUB ==")
    sys.exit(0 if run('manual push') else 1)
