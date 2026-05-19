from flask import (Flask, request, jsonify, render_template,
                   send_file, redirect, url_for, session, flash, g)
from flask_sqlalchemy import SQLAlchemy
from flask_login import (LoginManager, UserMixin, login_user, logout_user,
                         login_required, current_user)
from flask_wtf.csrf import CSRFProtect, generate_csrf
from flask_limiter import Limiter
from flask_limiter.util import get_remote_address
from werkzeug.security import generate_password_hash, check_password_hash
from dotenv import load_dotenv
import os, io, random, string
from analytics import process_file, generate_sample_template
from datetime import datetime, timedelta
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

load_dotenv()

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024
_db_url = os.getenv('DATABASE_URL', 'sqlite:///datapulse.db')
if _db_url.startswith('postgres://'):
    _db_url = _db_url.replace('postgres://', 'postgresql://', 1)
app.config['SQLALCHEMY_DATABASE_URI'] = _db_url
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
app.config['SECRET_KEY'] = os.getenv('SECRET_KEY', 'dev-secret-key-change-in-prod')
app.config['WTF_CSRF_ENABLED'] = True
app.config['WTF_CSRF_TIME_LIMIT'] = 3600
app.config['SESSION_COOKIE_HTTPONLY'] = True
app.config['SESSION_COOKIE_SAMESITE'] = 'Lax'

db = SQLAlchemy(app)
csrf = CSRFProtect(app)
login_manager = LoginManager(app)
login_manager.login_view = 'login_page'
login_manager.login_message = 'Connectez-vous pour accéder à cette page.'

limiter = Limiter(
    key_func=get_remote_address, app=app,
    default_limits=["200 per hour"], storage_uri="memory://"
)

ALLOWED_EXTENSIONS = {'xlsx', 'csv', 'xls'}
UPLOAD_FOLDER = 'uploaded_files'

# ── Models ────────────────────────────────────────────────────────────────────

class User(UserMixin, db.Model):
    __tablename__ = 'users'
    id            = db.Column(db.Integer, primary_key=True)
    email         = db.Column(db.String(255), unique=True, nullable=False)
    username      = db.Column(db.String(100), unique=True, nullable=False)
    password_hash = db.Column(db.String(512), nullable=False)
    created_at    = db.Column(db.DateTime, default=datetime.utcnow)
    imports       = db.relationship('Import', backref='owner', lazy='dynamic',
                                    cascade='all, delete-orphan')

    def set_password(self, password):
        self.password_hash = generate_password_hash(password)

    def check_password(self, password):
        return check_password_hash(self.password_hash, password)


class Import(db.Model):
    __tablename__ = 'imports'
    id              = db.Column(db.Integer, primary_key=True)
    user_id         = db.Column(db.Integer, db.ForeignKey('users.id'), nullable=False)
    original_name   = db.Column(db.String(255), nullable=False)
    display_name    = db.Column(db.String(255))
    file_extension  = db.Column(db.String(10),  nullable=False)
    file_size_bytes = db.Column(db.Integer)
    file_path       = db.Column(db.String(500))
    total_rows      = db.Column(db.Integer)
    valid_rows      = db.Column(db.Integer)
    status          = db.Column(db.String(20), default='pending')
    error_message   = db.Column(db.Text)
    imported_at     = db.Column(db.DateTime, default=datetime.utcnow)
    kpi             = db.relationship('KpiSnapshot', backref='import_ref',
                                      uselist=False, cascade='all, delete-orphan')
    status_agg      = db.relationship('AggStatus',   backref='import_ref',
                                      cascade='all, delete-orphan')
    trend_agg       = db.relationship('AggTrend',    backref='import_ref',
                                      cascade='all, delete-orphan')

    @property
    def label(self):
        return self.display_name or self.original_name


class KpiSnapshot(db.Model):
    __tablename__ = 'kpi_snapshots'
    id                = db.Column(db.Integer, primary_key=True)
    import_id         = db.Column(db.Integer, db.ForeignKey('imports.id'), unique=True)
    chiffre_affaires  = db.Column(db.Float)
    total_commandes   = db.Column(db.Integer)
    commandes_livrees = db.Column(db.Integer)
    panier_moyen      = db.Column(db.Float)
    taux_livraison    = db.Column(db.Float)
    date_min          = db.Column(db.String(20))
    date_max          = db.Column(db.String(20))
    computed_at       = db.Column(db.DateTime, default=datetime.utcnow)


class AggStatus(db.Model):
    __tablename__ = 'agg_status'
    id          = db.Column(db.Integer, primary_key=True)
    import_id   = db.Column(db.Integer, db.ForeignKey('imports.id'))
    label       = db.Column(db.String(50))
    order_count = db.Column(db.Integer)


class AggTrend(db.Model):
    __tablename__ = 'agg_trend'
    id        = db.Column(db.Integer, primary_key=True)
    import_id = db.Column(db.Integer, db.ForeignKey('imports.id'))
    day       = db.Column(db.String(20))
    revenue   = db.Column(db.Float)


class PendingUser(db.Model):
    __tablename__ = 'pending_users'
    id            = db.Column(db.Integer, primary_key=True)
    email         = db.Column(db.String(255), unique=True, nullable=False)
    username      = db.Column(db.String(100), nullable=False)
    password_hash = db.Column(db.String(512), nullable=False)
    code          = db.Column(db.String(6), nullable=False)
    expires_at    = db.Column(db.DateTime, nullable=False)


# ── Email helper ──────────────────────────────────────────────────────────────

def send_verification_email(to_email, code):
    gmail_user     = os.getenv('GMAIL_USER')
    gmail_password = os.getenv('GMAIL_APP_PASSWORD')
    if not gmail_user or not gmail_password:
        raise RuntimeError('GMAIL_USER or GMAIL_APP_PASSWORD not set in environment.')

    msg = MIMEMultipart('alternative')
    msg['Subject'] = 'DataPulse — Code de verification'
    msg['From']    = f'DataPulse <{gmail_user}>'
    msg['To']      = to_email

    html = f"""
    <div style="font-family:Inter,system-ui,sans-serif;background:#0f0e17;color:#e8e6f0;padding:40px;max-width:480px;margin:auto;border-radius:16px;">
      <div style="text-align:center;margin-bottom:24px;">
        <span style="font-size:2rem;">&#128202;</span>
        <h2 style="color:#818cf8;margin:8px 0 0;">DataPulse</h2>
      </div>
      <p style="margin-bottom:16px;">Bonjour <b>{to_email}</b>,</p>
      <p style="margin-bottom:24px;">Votre code de verification est :</p>
      <div style="background:#1a1825;border:1px solid #2d2a45;border-radius:12px;padding:24px;text-align:center;margin-bottom:24px;">
        <span style="font-size:2.5rem;font-weight:700;letter-spacing:12px;color:#818cf8;">{code}</span>
      </div>
      <p style="color:#8b89a0;font-size:.85rem;">Ce code expire dans <b>10 minutes</b>. Ne le partagez pas.</p>
    </div>
    """
    msg.attach(MIMEText(html, 'html'))

    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
        smtp.login(gmail_user, gmail_password)
        smtp.sendmail(gmail_user, to_email, msg.as_string())


# ── Auth setup ────────────────────────────────────────────────────────────────

@login_manager.user_loader
def load_user(user_id):
    return User.query.get(int(user_id))

def own_import_or_404(import_id):
    return Import.query.filter_by(id=import_id, user_id=current_user.id).first_or_404()

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

# ── Auth Routes ───────────────────────────────────────────────────────────────

@app.route('/login', methods=['GET'])
def login_page():
    if current_user.is_authenticated:
        return redirect(url_for('index'))
    return render_template('auth.html', mode='login')

@app.route('/register', methods=['GET'])
def register_page():
    if current_user.is_authenticated:
        return redirect(url_for('index'))
    return render_template('auth.html', mode='register')

@app.route('/api/login', methods=['POST'])
@limiter.limit("10 per minute")
def api_login():
    data = request.get_json() or {}
    email = data.get('email', '').strip().lower()
    password = data.get('password', '')
    user = User.query.filter_by(email=email).first()
    if not user or not user.check_password(password):
        return jsonify({'error': 'Email ou mot de passe incorrect.'}), 401
    login_user(user, remember=True)
    return jsonify({'ok': True, 'username': user.username})

@app.route('/api/register', methods=['POST'])
@limiter.limit("5 per minute")
def api_register():
    data = request.get_json() or {}
    email    = data.get('email', '').strip().lower()
    username = data.get('username', '').strip()
    password = data.get('password', '')
    if not email or not username or not password:
        return jsonify({'error': 'Tous les champs sont requis.'}), 400
    if len(password) < 6:
        return jsonify({'error': 'Le mot de passe doit contenir au moins 6 caracteres.'}), 400
    if User.query.filter_by(email=email).first():
        return jsonify({'error': 'Cet email est deja utilise.'}), 409
    if User.query.filter_by(username=username).first():
        return jsonify({'error': "Ce nom d utilisateur est deja pris."}), 409

    code = ''.join(random.choices(string.digits, k=6))
    expires = datetime.utcnow() + timedelta(minutes=10)
    PendingUser.query.filter_by(email=email).delete()
    pending = PendingUser(email=email, username=username, code=code, expires_at=expires)
    pending.password_hash = generate_password_hash(password)
    db.session.add(pending)
    db.session.commit()

    try:
        send_verification_email(email, code)
    except Exception as e:
        return jsonify({'error': f'Impossible d envoyer l email: {str(e)}'}), 500

    return jsonify({'ok': True, 'verify': True, 'email': email})


@app.route('/verify')
def verify_page():
    email = request.args.get('email', '')
    return render_template('verify.html', email=email)


@app.route('/api/verify', methods=['POST'])
@limiter.limit("10 per minute")
def api_verify():
    data  = request.get_json() or {}
    email = data.get('email', '').strip().lower()
    code  = data.get('code', '').strip()
    pending = PendingUser.query.filter_by(email=email).first()
    if not pending:
        return jsonify({'error': 'Aucune inscription en attente pour cet email.'}), 404
    if datetime.utcnow() > pending.expires_at:
        db.session.delete(pending)
        db.session.commit()
        return jsonify({'error': 'Code expire. Veuillez vous reinscrire.'}), 410
    if pending.code != code:
        return jsonify({'error': 'Code incorrect.'}), 401
    user = User(email=pending.email, username=pending.username)
    user.password_hash = pending.password_hash
    db.session.add(user)
    db.session.delete(pending)
    db.session.commit()
    login_user(user, remember=True)
    return jsonify({'ok': True, 'username': user.username})


@app.route('/api/resend-code', methods=['POST'])
@limiter.limit("3 per minute")
def api_resend_code():
    data  = request.get_json() or {}
    email = data.get('email', '').strip().lower()
    pending = PendingUser.query.filter_by(email=email).first()
    if not pending:
        return jsonify({'error': 'Aucune inscription en attente.'}), 404
    code = ''.join(random.choices(string.digits, k=6))
    pending.code       = code
    pending.expires_at = datetime.utcnow() + timedelta(minutes=10)
    db.session.commit()
    try:
        send_verification_email(email, code)
    except Exception as e:
        return jsonify({'error': f'Impossible d envoyer l email: {str(e)}'}), 500
    return jsonify({'ok': True})

@app.route('/logout')
@login_required
def logout():
    logout_user()
    return redirect(url_for('login_page'))

# ── App Routes ────────────────────────────────────────────────────────────────

@app.route('/')
@login_required
def index():
    return render_template('index.html', username=current_user.username)

@app.route('/dashboard')
@login_required
def dashboard():
    return render_template('dashboard.html', username=current_user.username)

@app.route('/history')
@login_required
def history():
    page      = request.args.get('page', 1, type=int)
    per_page  = 10
    date_from = request.args.get('date_from')
    date_to   = request.args.get('date_to')

    q = Import.query.filter_by(user_id=current_user.id)
    if date_from:
        try:
            q = q.filter(Import.imported_at >= datetime.strptime(date_from, '%Y-%m-%d'))
        except ValueError: pass
    if date_to:
        try:
            q = q.filter(Import.imported_at <= datetime.strptime(date_to, '%Y-%m-%d'))
        except ValueError: pass

    pagination = q.order_by(Import.imported_at.desc()).paginate(
        page=page, per_page=per_page, error_out=False)

    result = []
    for imp in pagination.items:
        entry = {
            'id':          imp.id,
            'name':        imp.label,
            'original':    imp.original_name,
            'status':      imp.status,
            'total_rows':  imp.total_rows,
            'imported_at': imp.imported_at.strftime('%d/%m/%Y %H:%M') if imp.imported_at else '',
            'has_file':    imp.file_path is not None and os.path.exists(imp.file_path),
        }
        if imp.kpi:
            entry['chiffre_affaires'] = imp.kpi.chiffre_affaires
            entry['taux_livraison']   = imp.kpi.taux_livraison
        result.append(entry)

    return jsonify({
        'items':    result,
        'total':    pagination.total,
        'pages':    pagination.pages,
        'current':  pagination.page,
        'has_next': pagination.has_next,
        'has_prev': pagination.has_prev,
    })

@app.route('/import/<int:import_id>')
@login_required
def get_import(import_id):
    imp = own_import_or_404(import_id)
    if not imp.kpi:
        return jsonify({'error': 'Données non disponibles'}), 404
    kpi = imp.kpi
    status_data = {
        'labels': [s.label       for s in imp.status_agg],
        'values': [s.order_count for s in imp.status_agg],
    }
    trend_data = {
        'labels': [t.day     for t in sorted(imp.trend_agg, key=lambda x: x.day)],
        'values': [t.revenue for t in sorted(imp.trend_agg, key=lambda x: x.day)],
    }
    date_from = request.args.get('date_from')
    date_to   = request.args.get('date_to')
    if date_from or date_to:
        pairs = list(zip(trend_data['labels'], trend_data['values']))
        if date_from:
            pairs = [(l, v) for l, v in pairs if l >= date_from]
        if date_to:
            pairs = [(l, v) for l, v in pairs if l <= date_to]
        trend_data = {'labels': [l for l,v in pairs], 'values': [v for l,v in pairs]}

    return jsonify({
        'kpis': {
            'chiffre_affaires': kpi.chiffre_affaires,
            'total_commandes':  kpi.total_commandes,
            'panier_moyen':     kpi.panier_moyen,
            'taux_livraison':   kpi.taux_livraison,
        },
        'status_breakdown': status_data,
        'trend_data':       trend_data,
        'total_rows':       imp.total_rows,
        'import_name':      imp.label,
        'import_date':      imp.imported_at.strftime('%d/%m/%Y %H:%M') if imp.imported_at else '',
        'has_file':         imp.file_path is not None and os.path.exists(imp.file_path),
    })

@app.route('/import/<int:import_id>', methods=['DELETE'])
@login_required
def delete_import(import_id):
    imp = own_import_or_404(import_id)
    if imp.file_path and os.path.exists(imp.file_path):
        try: os.remove(imp.file_path)
        except OSError: pass
    db.session.delete(imp)
    db.session.commit()
    return jsonify({'ok': True})

@app.route('/import/<int:import_id>/rename', methods=['PATCH'])
@login_required
def rename_import(import_id):
    imp  = own_import_or_404(import_id)
    data = request.get_json() or {}
    new_name = data.get('name', '').strip()
    if not new_name:
        return jsonify({'error': 'Nom invalide.'}), 400
    imp.display_name = new_name
    db.session.commit()
    return jsonify({'ok': True, 'name': new_name})

@app.route('/download/<int:import_id>')
@login_required
def download_file(import_id):
    imp = own_import_or_404(import_id)
    if not imp.file_path or not os.path.exists(imp.file_path):
        return jsonify({'error': 'Fichier original non disponible'}), 404
    mime_map = {
        'xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'xls':  'application/vnd.ms-excel',
        'csv':  'text/csv',
    }
    mime = mime_map.get(imp.file_extension, 'application/octet-stream')
    return send_file(imp.file_path, mimetype=mime, as_attachment=True, download_name=imp.original_name)

# ── Export routes ─────────────────────────────────────────────────────────────

@app.route('/export/excel/<int:import_id>')
@login_required
def export_excel(import_id):
    import pandas as pd
    imp = own_import_or_404(import_id)
    if not imp.kpi:
        return jsonify({'error': 'Données non disponibles'}), 404
    kpi = imp.kpi
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        pd.DataFrame([
            {'Indicateur': "Chiffre d'affaires (MAD)", 'Valeur': kpi.chiffre_affaires},
            {'Indicateur': 'Total commandes',           'Valeur': kpi.total_commandes},
            {'Indicateur': 'Panier moyen (MAD)',        'Valeur': kpi.panier_moyen},
            {'Indicateur': 'Taux de livraison (%)',     'Valeur': kpi.taux_livraison},
        ]).to_excel(writer, sheet_name='KPIs', index=False)

        pd.DataFrame({
            'Date': [t.day     for t in sorted(imp.trend_agg, key=lambda x: x.day)],
            'CA Livré (MAD)': [t.revenue for t in sorted(imp.trend_agg, key=lambda x: x.day)],
        }).to_excel(writer, sheet_name='Tendance', index=False)

        pd.DataFrame({
            'Statut':    [s.label       for s in imp.status_agg],
            'Commandes': [s.order_count for s in imp.status_agg],
        }).to_excel(writer, sheet_name='Statuts', index=False)

    output.seek(0)
    filename = f"analyse_{imp.label.replace(' ', '_')}.xlsx"
    return send_file(output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True, download_name=filename)

@app.route('/export/pdf/<int:import_id>')
@login_required
def export_pdf(import_id):
    try:
        from reportlab.lib.pagesizes import A4
        from reportlab.lib import colors
        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
        from reportlab.lib.units import cm
        from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
    except ImportError:
        return jsonify({'error': 'reportlab non installé. Lancez: pip install reportlab'}), 500

    imp = own_import_or_404(import_id)
    if not imp.kpi:
        return jsonify({'error': 'Données non disponibles'}), 404
    kpi = imp.kpi

    output = io.BytesIO()
    doc = SimpleDocTemplate(output, pagesize=A4,
                            rightMargin=2*cm, leftMargin=2*cm,
                            topMargin=2*cm, bottomMargin=2*cm)
    styles = getSampleStyleSheet()
    story = []

    title_style = ParagraphStyle('Title2', parent=styles['Title'],
                                 textColor=colors.HexColor('#4f46e5'), fontSize=20)
    story.append(Paragraph("Rapport d'analyse — DataPulse", title_style))
    story.append(Spacer(1, 0.4*cm))
    story.append(Paragraph(f"Import : <b>{imp.label}</b>", styles['Normal']))
    story.append(Paragraph(f"Généré le : {datetime.utcnow().strftime('%d/%m/%Y à %H:%M')}", styles['Normal']))
    story.append(Spacer(1, 0.8*cm))

    def make_table(data, header_color):
        t = Table(data, colWidths=[9*cm, 7*cm])
        t.setStyle(TableStyle([
            ('BACKGROUND',    (0,0), (-1,0), colors.HexColor(header_color)),
            ('TEXTCOLOR',     (0,0), (-1,0), colors.white),
            ('FONTNAME',      (0,0), (-1,0), 'Helvetica-Bold'),
            ('ROWBACKGROUNDS',(0,1), (-1,-1), [colors.HexColor('#f8f8ff'), colors.white]),
            ('GRID',          (0,0), (-1,-1), 0.5, colors.HexColor('#d1d5db')),
            ('ALIGN',         (1,0), (1,-1), 'RIGHT'),
            ('LEFTPADDING',   (0,0), (-1,-1), 8),
            ('RIGHTPADDING',  (0,0), (-1,-1), 8),
            ('TOPPADDING',    (0,0), (-1,-1), 6),
            ('BOTTOMPADDING', (0,0), (-1,-1), 6),
        ]))
        return t

    story.append(Paragraph('Indicateurs clés de performance', styles['Heading2']))
    story.append(Spacer(1, 0.3*cm))
    story.append(make_table([
        ['Indicateur', 'Valeur'],
        ["Chiffre d'affaires", f"{kpi.chiffre_affaires:,.2f} MAD"],
        ['Total commandes',    str(kpi.total_commandes)],
        ['Panier moyen',       f"{kpi.panier_moyen:,.2f} MAD"],
        ['Taux de livraison',  f"{kpi.taux_livraison:.1f} %"],
    ], '#4f46e5'))
    story.append(Spacer(1, 0.8*cm))

    story.append(Paragraph('Répartition par statut', styles['Heading2']))
    story.append(Spacer(1, 0.3*cm))
    story.append(make_table(
        [['Statut', 'Commandes']] + [[s.label, str(s.order_count)] for s in imp.status_agg],
        '#6366f1'
    ))

    doc.build(story)
    output.seek(0)
    filename = f"rapport_{imp.label.replace(' ', '_')}.pdf"
    return send_file(output, mimetype='application/pdf',
                     as_attachment=True, download_name=filename)

# ── Upload ────────────────────────────────────────────────────────────────────

@app.route('/upload', methods=['POST'])
@login_required
@limiter.limit("30 per hour")
def upload_file():
    if 'file' not in request.files:
        return jsonify({'error': 'Aucun fichier reçu.'}), 400
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'Nom de fichier vide.'}), 400
    if not allowed_file(file.filename):
        return jsonify({'error': 'Format non supporté. Utilisez .xlsx, .xls ou .csv'}), 400

    imp = Import(
        user_id=current_user.id,
        original_name=file.filename,
        file_extension=file.filename.rsplit('.', 1)[1].lower(),
        file_size_bytes=0,
        status='processing',
    )
    db.session.add(imp)
    db.session.flush()

    try:
        file_bytes = file.read()
        if len(file_bytes) == 0:
            return jsonify({'error': 'Le fichier est vide.'}), 422

        imp.file_size_bytes = len(file_bytes)
        os.makedirs(UPLOAD_FOLDER, exist_ok=True)
        safe_name  = f"{imp.id}_{file.filename.replace(' ', '_')}"
        saved_path = os.path.join(UPLOAD_FOLDER, safe_name)
        with open(saved_path, 'wb') as f:
            f.write(file_bytes)
        imp.file_path = saved_path

        result = process_file(file_bytes, imp.file_extension)

        kpi = KpiSnapshot(
            import_id         = imp.id,
            chiffre_affaires  = result['kpis']['chiffre_affaires'],
            total_commandes   = result['kpis']['total_commandes'],
            commandes_livrees = int(result['kpis']['total_commandes'] * result['kpis']['taux_livraison'] / 100),
            panier_moyen      = result['kpis']['panier_moyen'],
            taux_livraison    = result['kpis']['taux_livraison'],
            date_min          = result['trend_data']['labels'][0]  if result['trend_data']['labels'] else None,
            date_max          = result['trend_data']['labels'][-1] if result['trend_data']['labels'] else None,
        )
        db.session.add(kpi)

        for label, count in zip(result['status_breakdown']['labels'], result['status_breakdown']['values']):
            db.session.add(AggStatus(import_id=imp.id, label=label, order_count=count))
        for day, rev in zip(result['trend_data']['labels'], result['trend_data']['values']):
            db.session.add(AggTrend(import_id=imp.id, day=day, revenue=rev))

        imp.status     = 'success'
        imp.total_rows = result['total_rows']
        imp.valid_rows = result['total_rows']
        db.session.commit()

        return jsonify({
            'import_id':        imp.id,
            'kpis':             result['kpis'],
            'status_breakdown': result['status_breakdown'],
            'trend_data':       result['trend_data'],
            'total_rows':       result['total_rows'],
            'import_name':      imp.label,
            'import_date':      imp.imported_at.strftime('%d/%m/%Y %H:%M'),
            'has_file':         True,
        })

    except ValueError as e:
        db.session.rollback()
        imp.status = 'error'; imp.error_message = str(e); db.session.commit()
        return jsonify({'error': str(e)}), 422
    except Exception as e:
        db.session.rollback()
        imp.status = 'error'; imp.error_message = str(e); db.session.commit()
        return jsonify({'error': f'Erreur de traitement: {str(e)}'}), 500

# ── Multi-file upload (batch) ─────────────────────────────────────────────────

@app.route('/upload/multi', methods=['POST'])
@login_required
@limiter.limit("20 per hour")
def upload_multi():
    files = request.files.getlist('files')
    if not files or all(f.filename == '' for f in files):
        return jsonify({'error': 'Aucun fichier reçu.'}), 400
    results = []
    for file in files:
        if file.filename == '' or not allowed_file(file.filename):
            results.append({'name': file.filename, 'error': 'Format non supporté'}); continue
        imp = Import(user_id=current_user.id, original_name=file.filename,
                     file_extension=file.filename.rsplit('.', 1)[1].lower(),
                     file_size_bytes=0, status='processing')
        db.session.add(imp); db.session.flush()
        try:
            file_bytes = file.read()
            if not file_bytes:
                results.append({'name': file.filename, 'error': 'Fichier vide'}); continue
            imp.file_size_bytes = len(file_bytes)
            os.makedirs(UPLOAD_FOLDER, exist_ok=True)
            saved_path = os.path.join(UPLOAD_FOLDER, f"{imp.id}_{file.filename.replace(' ','_')}")
            with open(saved_path, 'wb') as fh: fh.write(file_bytes)
            imp.file_path = saved_path
            result = process_file(file_bytes, imp.file_extension)
            kpi = KpiSnapshot(
                import_id=imp.id,
                chiffre_affaires=result['kpis']['chiffre_affaires'],
                total_commandes=result['kpis']['total_commandes'],
                commandes_livrees=int(result['kpis']['total_commandes'] * result['kpis']['taux_livraison'] / 100),
                panier_moyen=result['kpis']['panier_moyen'],
                taux_livraison=result['kpis']['taux_livraison'],
                date_min=result['trend_data']['labels'][0]  if result['trend_data']['labels'] else None,
                date_max=result['trend_data']['labels'][-1] if result['trend_data']['labels'] else None,
            )
            db.session.add(kpi)
            for label, count in zip(result['status_breakdown']['labels'], result['status_breakdown']['values']):
                db.session.add(AggStatus(import_id=imp.id, label=label, order_count=count))
            for day, rev in zip(result['trend_data']['labels'], result['trend_data']['values']):
                db.session.add(AggTrend(import_id=imp.id, day=day, revenue=rev))
            imp.status = 'success'; imp.total_rows = result['total_rows']; imp.valid_rows = result['total_rows']
            db.session.commit()
            results.append({'name': file.filename, 'import_id': imp.id,
                            'kpis': result['kpis'], 'total_rows': result['total_rows']})
        except Exception as e:
            db.session.rollback()
            imp.status = 'error'; imp.error_message = str(e); db.session.commit()
            results.append({'name': file.filename, 'error': str(e)})
    return jsonify({'results': results})


# ── Merge imports ─────────────────────────────────────────────────────────────

@app.route('/merge', methods=['POST'])
@login_required
def merge_imports():
    data = request.get_json() or {}
    ids  = data.get('import_ids', [])
    if len(ids) < 2:
        return jsonify({'error': 'Sélectionnez au moins 2 imports à fusionner.'}), 400
    imports = Import.query.filter(Import.id.in_(ids), Import.user_id == current_user.id).all()
    if len(imports) < 2:
        return jsonify({'error': 'Imports introuvables.'}), 404
    total_ca  = sum(i.kpi.chiffre_affaires  for i in imports if i.kpi) or 0
    total_cmd = sum(i.kpi.total_commandes   for i in imports if i.kpi) or 0
    total_liv = sum(i.kpi.commandes_livrees for i in imports if i.kpi) or 0
    panier_moyen   = round(total_ca / total_liv, 2) if total_liv else 0
    taux_livraison = round(total_liv / total_cmd * 100, 1) if total_cmd else 0
    trend_map = {}
    for imp in imports:
        for t in imp.trend_agg:
            trend_map[t.day] = trend_map.get(t.day, 0) + t.revenue
    trend_sorted = sorted(trend_map.items())
    status_map = {}
    for imp in imports:
        for s in imp.status_agg:
            status_map[s.label] = status_map.get(s.label, 0) + s.order_count
    return jsonify({
        'merged': True,
        'import_names': [i.label for i in imports],
        'kpis': {'chiffre_affaires': round(total_ca, 2), 'total_commandes': total_cmd,
                 'panier_moyen': panier_moyen, 'taux_livraison': taux_livraison},
        'trend_data':       {'labels': [d for d, _ in trend_sorted], 'values': [round(v,2) for _,v in trend_sorted]},
        'status_breakdown': {'labels': list(status_map.keys()), 'values': list(status_map.values())},
        'total_rows': sum(i.total_rows or 0 for i in imports),
    })


# ── Growth KPI ────────────────────────────────────────────────────────────────

@app.route('/import/<int:import_id>/growth')
@login_required
def get_growth(import_id):
    current = own_import_or_404(import_id)
    if not current.kpi:
        return jsonify({'error': 'Données non disponibles'}), 404
    prev = Import.query.filter(
        Import.user_id == current_user.id, Import.id != import_id,
        Import.status == 'success', Import.imported_at < current.imported_at,
    ).order_by(Import.imported_at.desc()).first()
    if not prev or not prev.kpi:
        return jsonify({'has_prev': False})
    def pct(new, old): return round((new - old) / old * 100, 1) if old else None
    return jsonify({
        'has_prev': True, 'prev_name': prev.label,
        'growth': {
            'chiffre_affaires': pct(current.kpi.chiffre_affaires, prev.kpi.chiffre_affaires),
            'total_commandes':  pct(current.kpi.total_commandes,  prev.kpi.total_commandes),
            'panier_moyen':     pct(current.kpi.panier_moyen,     prev.kpi.panier_moyen),
            'taux_livraison':   pct(current.kpi.taux_livraison,   prev.kpi.taux_livraison),
        },
        'alerts': {'low_delivery': current.kpi.taux_livraison < 70, 'taux_livraison': current.kpi.taux_livraison}
    })


# ── Monthly comparison ────────────────────────────────────────────────────────

@app.route('/compare/monthly')
@login_required
def compare_monthly():
    imports = Import.query.filter_by(user_id=current_user.id, status='success').order_by(Import.imported_at).all()
    months = {}
    for imp in imports:
        if not imp.kpi: continue
        mk = imp.imported_at.strftime('%m/%Y') if imp.imported_at else 'Inconnu'
        if mk not in months: months[mk] = {'ca':0,'cmd':0,'liv':0,'count':0}
        months[mk]['ca']  += imp.kpi.chiffre_affaires  or 0
        months[mk]['cmd'] += imp.kpi.total_commandes   or 0
        months[mk]['liv'] += imp.kpi.commandes_livrees or 0
        months[mk]['count'] += 1
    labels = list(months.keys())
    return jsonify({
        'labels': labels,
        'ca':   [round(months[m]['ca'],2) for m in labels],
        'cmd':  [months[m]['cmd'] for m in labels],
        'taux': [round(months[m]['liv']/months[m]['cmd']*100,1) if months[m]['cmd'] else 0 for m in labels],
    })


@app.route('/template')
def download_template():
    output = generate_sample_template()
    return send_file(output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True, download_name='template_ventes.xlsx')

@app.route('/api/me')
@login_required
def api_me():
    return jsonify({'username': current_user.username, 'email': current_user.email})

@app.after_request
def set_csrf_cookie(response):
    response.set_cookie('csrf_token', generate_csrf())
    return response

with app.app_context():
    db.create_all()

if __name__ == '__main__':
    app.run(debug=True, port=5000)
