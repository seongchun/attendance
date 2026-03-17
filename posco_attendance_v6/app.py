"""
광양제철소 행정섭외그룹 – QR 출석 관리 서버 (공용 QR + ngrok 터널링)
"""

import io, os, json, base64, socket, smtplib, logging, uuid
from datetime import datetime, timezone
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.header import Header
from email.utils import formatdate, make_msgid

from flask import Flask, render_template, request, jsonify, send_file
from apscheduler.schedulers.background import BackgroundScheduler
from apscheduler.triggers.cron import CronTrigger
import qrcode

logging.basicConfig(level=logging.INFO, format='%(asctime)s %(levelname)s %(message)s')
log = logging.getLogger(__name__)

app = Flask(__name__)
os.makedirs('data', exist_ok=True)

CONFIG_FILE = 'data/config.json'
DEFAULT_CONFIG = {
    "meeting_name":       "그룹 현안회의",
    "meeting_place":      "행정섭외그룹 회의실",
    "email_to":           "",
    "email_cc":           "",
    "smtp_host":          "smtp.gmail.com",
    "smtp_port":          587,
    "smtp_user":          "",
    "smtp_pass":          "",
    "sender_name":        "행정기획 담당자",
    "auto_send_enabled":  False,
    "auto_send_time":     "10:30",
    "email_method":       "smtp",
    "brevo_api_key":      "",
    "brevo_sender_email": "",
    "ngrok_authtoken":    "",
    "members": [
        {"id": "m01", "name": "홍길동",  "section": "그룹장"},
        {"id": "m02", "name": "김영수",  "section": "행정보안섹션"},
        {"id": "m03", "name": "이수진",  "section": "행정보안섹션"},
        {"id": "m04", "name": "박민준",  "section": "홍보섹션"},
        {"id": "m05", "name": "최지원",  "section": "홍보섹션"},
        {"id": "m06", "name": "정대한",  "section": "대외협력섹션"},
        {"id": "m07", "name": "한소희",  "section": "대외협력섹션"},
        {"id": "m08", "name": "윤재영",  "section": "후생섹션"},
        {"id": "m09", "name": "임재민",  "section": "후생섹션"},
        {"id": "m10", "name": "강나연",  "section": "후생섹션"},
    ]
}

attendance: dict = {}
public_url: str = ""


def _normalize_members(cfg: dict) -> dict:
    """section ↔ dept 하위호환 정규화 (기존 config.json 대응)"""
    for m in cfg.get('members', []):
        if 'section' not in m:
            m['section'] = m.get('dept', '')
        m['dept'] = m['section']   # 레거시 필드 동기화
    return cfg

def load_config() -> dict:
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, encoding='utf-8') as f:
                return _normalize_members(json.load(f))
        except Exception:
            pass
    return _normalize_members(DEFAULT_CONFIG.copy())

def save_config(cfg: dict):
    with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)

def get_local_ip() -> str:
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except Exception:
        return "127.0.0.1"

def get_checkin_base_url() -> str:
    if public_url:
        return public_url
    return f"http://{get_local_ip()}:5000"

def start_ngrok(port: int = 5000) -> str:
    """ngrok 터널 시도"""
    global public_url
    cfg = load_config()
    token = cfg.get('ngrok_authtoken', '').strip()
    if not token:
        return ""
    try:
        from pyngrok import ngrok, conf
        conf.get_default().auth_token = token
        tunnel = ngrok.connect(port, bind_tls=True)
        public_url = tunnel.public_url
        log.info('ngrok tunnel: %s', public_url)
        return public_url
    except ImportError:
        log.warning('pyngrok not installed')
        return ""
    except Exception as e:
        log.warning('ngrok failed: %s', e)
        return ""


_cf_process = None   # cloudflared 프로세스 참조

def start_cloudflare_tunnel(port: int = 5000) -> str:
    """Cloudflare Quick Tunnel (계정 불필요, 기업 방화벽 우회 가능성 높음)"""
    global public_url, _cf_process
    import subprocess, sys, platform, re, time, urllib.request, zipfile, stat

    # cloudflared 바이너리 경로 결정
    is_win = platform.system() == 'Windows'
    cf_bin = os.path.join('data', 'cloudflared.exe' if is_win else 'cloudflared')

    # 바이너리가 없으면 다운로드
    if not os.path.exists(cf_bin):
        log.info('Downloading cloudflared ...')
        try:
            if is_win:
                url = 'https://github.com/cloudflare/cloudflared/releases/latest/download/cloudflared-windows-amd64.exe'
                urllib.request.urlretrieve(url, cf_bin)
            else:
                arch = platform.machine()
                fname = 'cloudflared-linux-amd64' if 'x86_64' in arch else 'cloudflared-linux-arm64'
                url = f'https://github.com/cloudflare/cloudflared/releases/latest/download/{fname}'
                urllib.request.urlretrieve(url, cf_bin)
                st = os.stat(cf_bin)
                os.chmod(cf_bin, st.st_mode | stat.S_IEXEC)
            log.info('cloudflared downloaded: %s', cf_bin)
        except Exception as e:
            log.warning('cloudflared download failed: %s', e)
            return ""

    # cloudflared 실행
    try:
        _cf_process = subprocess.Popen(
            [cf_bin, 'tunnel', '--url', f'http://localhost:{port}'],
            stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
            text=True, encoding='utf-8', errors='replace'
        )
        # 출력에서 trycloudflare.com URL 추출 (최대 30초 대기)
        deadline = time.time() + 30
        while time.time() < deadline:
            line = _cf_process.stdout.readline()
            if not line:
                time.sleep(0.5)
                continue
            match = re.search(r'(https://[a-zA-Z0-9\-]+\.trycloudflare\.com)', line)
            if match:
                public_url = match.group(1)
                log.info('Cloudflare tunnel: %s', public_url)
                return public_url
        log.warning('cloudflare tunnel: URL not found in output')
        return ""
    except Exception as e:
        log.warning('cloudflare tunnel failed: %s', e)
        return ""


def start_tunnel(port: int = 5000) -> str:
    """터널 시도: ngrok -> cloudflare -> 로컬네트워크 안내"""
    # 1) ngrok 시도
    url = start_ngrok(port)
    if url:
        return url
    # 2) Cloudflare Quick Tunnel 시도
    log.info('ngrok unavailable, trying Cloudflare tunnel...')
    url = start_cloudflare_tunnel(port)
    if url:
        return url
    # 3) 터널 없음 - 로컬 네트워크만 사용
    log.info('No tunnel available - local network only')
    return ""

def make_qr_png_bytes(url: str) -> bytes:
    qr = qrcode.QRCode(version=None,
                       error_correction=qrcode.constants.ERROR_CORRECT_M,
                       box_size=10, border=3)
    qr.add_data(url)
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white")
    buf = io.BytesIO()
    img.save(buf, format='PNG')
    return buf.getvalue()

def build_report_body(cfg: dict) -> str:
    members   = cfg.get('members', [])
    present   = [m for m in members if m['id'] in attendance]
    absent    = [m for m in members if m['id'] not in attendance]
    total     = len(members)
    rate      = round(len(present) / total * 100) if total else 0
    now_str   = datetime.now().strftime('%Y년 %m월 %d일 %H:%M')
    sender    = cfg.get('sender_name', '담당자')
    mtg_name  = cfg.get('meeting_name', '그룹 현안회의')
    mtg_place = cfg.get('meeting_place', '')
    lines = [
        f"안녕하십니까, {sender}입니다.\n",
        f"{mtg_name} 출석 현황을 아래와 같이 보고드립니다.\n",
        "-" * 40,
        f" 회 의 명 : {mtg_name}",
        f" 보고일시 : {now_str}",
        f" 장   소 : {mtg_place}",
        f" 참석 현황: {len(present)}명 / 전체 {total}명  (출석률 {rate}%)",
        "-" * 40 + "\n",
        f"[ 출석자 ({len(present)}명) ]",
    ]
    if present:
        for i, m in enumerate(present, 1):
            t = attendance[m['id']]['time']
            lines.append(f"  {i}. {m['name']} ({m.get('section', m.get('dept', ''))}) - QR 체크인: {t}")
    else:
        lines.append("  없음")
    lines.append(f"\n[ 미출석 ({len(absent)}명) ]")
    if absent:
        for i, m in enumerate(absent, 1):
            lines.append(f"  {i}. {m['name']} ({m.get('section', m.get('dept', ''))})")
    else:
        lines.append("  없음 (전원 출석)")
    lines += ["\n" + "-" * 40,
              f"감사합니다.\n광양제철소 행정섭외그룹  {sender} 드림"]
    return '\n'.join(lines)

def _build_subject(cfg: dict) -> str:
    members = cfg.get('members', [])
    present_cnt = sum(1 for m in members if m['id'] in attendance)
    total = len(members)
    now = datetime.now()
    return (f"[{cfg.get('meeting_name','현안회의')} "
            f"{now.strftime('%Y.%m.%d')}] 출석 현황 보고 - {present_cnt}/{total}명")


def send_email_brevo(cfg: dict) -> tuple[bool, str]:
    """Brevo(Sendinblue) HTTPS API로 메일 발송 (포트 443, 방화벽 우회)"""
    import urllib.request, urllib.error
    api_key = cfg.get('brevo_api_key', '').strip()
    sender_email = cfg.get('brevo_sender_email', '').strip()
    sender_name = cfg.get('sender_name', '담당자')
    to = cfg.get('email_to', '').strip()
    cc = cfg.get('email_cc', '').strip()

    if not api_key:
        return False, 'Brevo API 키가 설정되지 않았습니다.'
    if not sender_email:
        return False, 'Brevo 발신자 이메일이 설정되지 않았습니다.'
    if not to:
        return False, '수신자 이메일이 설정되지 않았습니다.'

    subject = _build_subject(cfg)
    body = build_report_body(cfg)

    payload = {
        "sender": {"name": sender_name, "email": sender_email},
        "to": [{"email": addr.strip()} for addr in to.split(',') if addr.strip()],
        "subject": subject,
        "textContent": body,
    }
    if cc:
        payload["cc"] = [{"email": addr.strip()} for addr in cc.split(',') if addr.strip()]

    data = json.dumps(payload).encode('utf-8')
    api_url = 'https://api.brevo.com/v3/smtp/email'
    hdrs = {
        'accept': 'application/json',
        'api-key': api_key,
        'content-type': 'application/json',
    }

    def _make_request(ssl_ctx=None):
        r = urllib.request.Request(api_url, data=data, headers=hdrs, method='POST')
        return urllib.request.urlopen(r, timeout=30, context=ssl_ctx)

    try:
        import ssl
        # 1차: 기본 SSL (일반 PC)
        try:
            with _make_request() as resp:
                log.info('Brevo email sent: %s', resp.read().decode())
                return True, '메일이 성공적으로 발송되었습니다. (Brevo API)'
        except (ssl.SSLError, urllib.error.URLError) as ssl_err:
            # 2차: 완화된 SSL (회사 프록시 환경)
            log.warning('Default SSL failed (%s), retrying with relaxed SSL...', ssl_err)
            ctx = ssl.SSLContext(ssl.PROTOCOL_TLS_CLIENT)
            ctx.check_hostname = False
            ctx.verify_mode = ssl.CERT_NONE
            ctx.set_ciphers('DEFAULT:@SECLEVEL=0')
            with _make_request(ssl_ctx=ctx) as resp:
                log.info('Brevo email sent (relaxed SSL): %s', resp.read().decode())
                return True, '메일이 성공적으로 발송되었습니다. (Brevo API)'
    except urllib.error.HTTPError as e:
        err_body = e.read().decode('utf-8', errors='replace')
        log.error('Brevo API error %s: %s', e.code, err_body)
        if e.code == 401:
            return False, f'Brevo API 인증 실패 (401). 키를 확인해 주세요.\n응답: {err_body}'
        return False, f'Brevo 발송 실패 (HTTP {e.code}): {err_body}'
    except Exception as e:
        log.error('Brevo error: %s', e)
        return False, f'Brevo 발송 실패: {str(e)}'


def send_email(cfg: dict = None) -> tuple[bool, str]:
    if cfg is None:
        cfg = load_config()

    method = cfg.get('email_method', 'smtp').strip().lower()

    # Brevo API 방식
    if method == 'brevo':
        return send_email_brevo(cfg)

    # 기존 SMTP 방식
    to = cfg.get('email_to', '').strip()
    cc = cfg.get('email_cc', '').strip()
    if not to:
        return False, '수신자 이메일이 설정되지 않았습니다.'
    subject = _build_subject(cfg)
    msg = MIMEMultipart()
    msg['From'] = cfg.get('smtp_user', 'noreply')
    msg['To'] = to
    if cc:
        msg['Cc'] = cc
    msg['Subject'] = Header(subject, 'utf-8')
    msg['Date'] = formatdate(localtime=True)
    sender_domain = (cfg.get('smtp_user', '') or 'local').split('@')[-1] or 'local'
    msg['Message-ID'] = make_msgid(domain=sender_domain)
    msg.attach(MIMEText(build_report_body(cfg), 'plain', 'utf-8'))
    host = cfg.get('smtp_host', '')
    port = int(cfg.get('smtp_port', 587))
    user = cfg.get('smtp_user', '').strip()
    pw   = cfg.get('smtp_pass', '').strip()
    recipients = [to] + ([cc] if cc else [])
    msg_bytes = msg.as_bytes()
    try:
        if port == 465:
            import ssl
            with smtplib.SMTP_SSL(host, port, context=ssl.create_default_context()) as s:
                if user and pw: s.login(user, pw)
                s.sendmail(user or 'noreply', recipients, msg_bytes)
        else:
            with smtplib.SMTP(host, port, timeout=15) as s:
                s.ehlo()
                if port == 587: s.starttls(); s.ehlo()
                if user and pw: s.login(user, pw)
                s.sendmail(user or 'noreply', recipients, msg_bytes)
        log.info('Email sent to %s', to)
        return True, '메일이 성공적으로 발송되었습니다.'
    except Exception as e:
        log.error('Email error: %s', e)
        err_str = str(e)
        # Gmail 인증 오류 → 앱 비밀번호 안내
        if '530' in err_str or '535' in err_str or 'Authentication' in err_str or 'Username and Password' in err_str:
            return False, (
                'Gmail 인증 오류: 일반 비밀번호는 사용할 수 없습니다.\n'
                '해결 방법 → Google 계정 보안 설정에서 [앱 비밀번호]를 발급받아 입력하세요.\n'
                '(설정 → 비밀번호 입력란에 16자리 앱 비밀번호 입력 후 저장)\n'
                f'원본 오류: {err_str}'
            )
        return False, err_str

scheduler = BackgroundScheduler(timezone='Asia/Seoul')

def setup_scheduler():
    scheduler.remove_all_jobs()
    cfg = load_config()
    if cfg.get('auto_send_enabled'):
        t = cfg.get('auto_send_time', '10:30')
        h, m = t.split(':')
        scheduler.add_job(lambda: send_email(load_config()),
                          CronTrigger(hour=int(h), minute=int(m)),
                          id='auto_email', replace_existing=True)
        log.info('Auto-email scheduled: %s', t)


@app.route('/')
def admin():
    cfg     = load_config()
    ip      = get_local_ip()
    members = cfg.get('members', [])
    present = sum(1 for m in members if m['id'] in attendance)
    base    = get_checkin_base_url()
    qr_url  = base + '/checkin'
    qr_b64  = base64.b64encode(make_qr_png_bytes(qr_url)).decode()
    local_qr_url = f"http://{ip}:5000/checkin"
    local_qr_b64 = base64.b64encode(make_qr_png_bytes(local_qr_url)).decode()
    return render_template('admin.html',
        config=cfg, members=members, attendance=attendance,
        ip=ip, port=5000,
        present_count=present, absent_count=len(members)-present, total=len(members),
        qr_url=qr_url, qr_b64=qr_b64, public_url=public_url, tunnel_active=bool(public_url),
        local_qr_url=local_qr_url, local_qr_b64=local_qr_b64)

@app.route('/qr')
def qr_sheet():
    cfg    = load_config()
    ip     = get_local_ip()
    base   = get_checkin_base_url()
    qr_url = base + '/checkin'
    qr_b64 = base64.b64encode(make_qr_png_bytes(qr_url)).decode()
    return render_template('qr_sheet.html',
        meeting_name=cfg.get('meeting_name', '그룹 현안회의'),
        ip=ip, qr_url=qr_url, qr_b64=qr_b64, tunnel_active=bool(public_url))

@app.route('/checkin')
def checkin_page():
    cfg     = load_config()
    members = cfg.get('members', [])
    return render_template('checkin.html',
        members=members, checked_ids=list(attendance.keys()),
        attendance=attendance, meeting_name=cfg.get('meeting_name', '그룹 현안회의'))

@app.route('/api/checkin', methods=['POST'])
def api_checkin():
    data      = request.get_json(force=True)
    member_id = data.get('id', '')
    cfg       = load_config()
    member    = next((m for m in cfg.get('members', []) if m['id'] == member_id), None)
    if not member:
        return jsonify({'ok': False, 'msg': '회원을 찾을 수 없습니다.'})
    if member_id in attendance:
        return jsonify({'ok': False, 'already': True,
                        'name': member['name'], 'time': attendance[member_id]['time']})
    now = datetime.now().strftime('%H:%M:%S')
    sec = member.get('section', member.get('dept', ''))
    attendance[member_id] = {'name': member['name'], 'section': sec, 'dept': sec, 'time': now}
    log.info('CHECK-IN: %s (%s) %s', member['name'], member['dept'], now)
    return jsonify({'ok': True, 'name': member['name'], 'time': now})

@app.route('/api/cancel', methods=['POST'])
def api_cancel():
    attendance.pop(request.get_json(force=True).get('id', ''), None)
    return jsonify({'ok': True})

@app.route('/api/status')
def api_status():
    cfg     = load_config()
    members = cfg.get('members', [])
    rows    = [{'id': m['id'], 'name': m['name'], 'dept': m['dept'],
                'present': bool(attendance.get(m['id'])),
                'time': attendance[m['id']]['time'] if m['id'] in attendance else None}
               for m in members]
    present = sum(1 for r in rows if r['present'])
    return jsonify({'members': rows, 'total': len(rows), 'present': present,
                    'absent': len(rows)-present,
                    'rate': round(present/len(rows)*100) if rows else 0})

@app.route('/api/reset', methods=['POST'])
def api_reset():
    attendance.clear()
    return jsonify({'ok': True})

@app.route('/api/send_email', methods=['POST'])
def api_send_email():
    ok, msg = send_email()
    return jsonify({'ok': ok, 'msg': msg})

@app.route('/api/report_data')
def api_report_data():
    """메일 앱 열기 / 클립보드 복사용 보고서 데이터"""
    cfg = load_config()
    members = cfg.get('members', [])
    present_cnt = sum(1 for m in members if m['id'] in attendance)
    total = len(members)
    now = datetime.now()
    subject = (f"[{cfg.get('meeting_name','현안회의')} "
               f"{now.strftime('%Y.%m.%d')}] 출석 현황 보고 - {present_cnt}/{total}명")
    body = build_report_body(cfg)
    return jsonify({
        'subject': subject,
        'body': body,
        'to': cfg.get('email_to', ''),
        'cc': cfg.get('email_cc', ''),
    })

@app.route('/api/report_download')
def api_report_download():
    """보고서 텍스트 파일 다운로드"""
    cfg = load_config()
    body = build_report_body(cfg)
    members = cfg.get('members', [])
    present_cnt = sum(1 for m in members if m['id'] in attendance)
    total = len(members)
    now = datetime.now()
    fname = f"출석보고_{now.strftime('%Y%m%d_%H%M')}_{present_cnt}of{total}.txt"
    buf = io.BytesIO(body.encode('utf-8'))
    return send_file(buf, mimetype='text/plain; charset=utf-8',
                     as_attachment=True, download_name=fname)

@app.route('/api/save_config', methods=['POST'])
def api_save_config():
    global public_url
    data      = request.get_json(force=True)
    old_cfg   = load_config()
    new_token = data.get('ngrok_authtoken', '').strip()
    old_token = old_cfg.get('ngrok_authtoken', '').strip()
    save_config(data)
    setup_scheduler()
    if new_token != old_token and new_token:
        try:
            from pyngrok import ngrok; ngrok.kill()
        except Exception:
            pass
        start_tunnel()
    return jsonify({'ok': True, 'public_url': public_url})

@app.route('/api/qr_image')
def api_qr_image():
    url = get_checkin_base_url() + '/checkin'
    return send_file(io.BytesIO(make_qr_png_bytes(url)), mimetype='image/png')


if __name__ == '__main__':
    setup_scheduler()
    if not scheduler.running:
        scheduler.start()
    ip = get_local_ip()
    tunnel_url = start_tunnel()
    print("=" * 56)
    print("  POSCO Gwangyang - QR Attendance Server")
    print("=" * 56)
    print(f"  Admin (PC) : http://localhost:5000")
    print(f"  Local IP   : http://{ip}:5000")
    if tunnel_url:
        print(f"  Public URL : {tunnel_url}")
        print(f"  QR Check-in: {tunnel_url}/checkin")
        print("  * Tunnel ACTIVE - cellular data phones can connect")
    else:
        print()
        print(f"  ** 터널 미활성 - 같은 WiFi 기기만 접속 가능 **")
        print(f"  ** 휴대폰을 회사 WiFi에 연결 후 QR 스캔하세요 **")
    print("  Stop       : Ctrl+C")
    print("=" * 56)
    app.run(host='0.0.0.0', port=5000, debug=False, use_reloader=False)
