from datetime import datetime
from functools import wraps
from io import BytesIO
import base64
import random
import string

from flask import Flask, flash, redirect, render_template, request, send_file, session, url_for
from flask_sqlalchemy import SQLAlchemy
from sqlalchemy import func, text
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import SimpleDocTemplate, Spacer, Table, TableStyle, Paragraph


app = Flask(__name__)
app.config["SECRET_KEY"] = "ticket-secret-key"
app.config["SQLALCHEMY_DATABASE_URI"] = "sqlite:///database.db"
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False
db = SQLAlchemy(app)

TRANSLATIONS = {
    "fr": {
        "dashboard": "Tableau de bord", "events": "Événements", "categories": "Catégories",
        "users": "Utilisateurs", "sales": "Ventes", "tickets": "Billets", "control": "Contrôle",
        "reports": "Rapports",
        "account": "Compte", "logout": "Déconnexion", "login": "Connexion", "language": "Langue",
        "currency": "Devise", "french": "Français", "english": "Anglais", "save": "Enregistrer",
        "settings_saved": "Préférences enregistrées.", "role": "Rôle", "exchange_rate": "Taux de conversion",
        "exchange_help": "Nombre de FC pour 1 dollar américain.", "fc": "FC", "usd": "$",
        "event_exchange_rate": "Taux de conversion de l'événement", "event_exchange_help": "1 $ = combien de FC ?",
        "Performance globale": "Performance globale", "Bienvenue": "Bienvenue", "Nouvelle vente": "Nouvelle vente",
        "Événements": "Événements", "Ventes": "Ventes", "Billets": "Billets", "Revenu total": "Revenu total",
        "Événements à venir": "Événements à venir", "Prix moyen billet": "Prix moyen billet",
        "Taux de remplissage": "Taux de remplissage", "Billets validés": "Billets validés",
        "Programmation active": "Programmation active", "Par vente": "Par vente", "Capacité moyenne": "Capacité moyenne",
        "Contrôle d’entrée": "Contrôle d’entrée", "Transactions": "Transactions", "Ventes récentes": "Ventes récentes",
        "Activité": "Activité", "Derniers billets générés": "Derniers billets générés", "Performance": "Performance",
        "Points clés": "Points clés", "Aucun événement pour le moment.": "Aucun événement pour le moment.",
        "Liste des événements": "Liste des événements", "Ajouter un événement": "Ajouter un événement",
        "Catégories et prix": "Catégories et prix", "Ajouter une catégorie": "Ajouter une catégorie",
        "Vente de billets": "Vente de billets", "Sélectionner": "Sélectionner", "Valider la vente": "Valider la vente",
        "Billets vendus": "Billets vendus", "Contrôle d’entrée": "Contrôle d’entrée", "Valider l’entrée": "Valider l’entrée",
        "Rapports de ventes": "Rapports de ventes", "Rapport journalier": "Rapport journalier", "Rapport mensuel": "Rapport mensuel",
        "Filtrer": "Filtrer", "Télécharger PDF": "Télécharger PDF", "Télécharger Excel": "Télécharger Excel",
        "Chiffre d’affaires": "Chiffre d’affaires", "Nombre de ventes": "Nombre de ventes", "Billets vendus": "Billets vendus",
    },
    "en": {
        "dashboard": "Dashboard", "events": "Events", "categories": "Categories", "users": "Users",
        "sales": "Sales", "tickets": "Tickets", "control": "Check-in", "account": "Account",
        "reports": "Reports",
        "logout": "Log out", "login": "Log in", "language": "Language", "currency": "Currency",
        "french": "French", "english": "English", "save": "Save", "settings_saved": "Preferences saved.",
        "role": "Role", "exchange_rate": "Exchange rate", "exchange_help": "Number of FC for 1 US dollar.",
        "fc": "FC", "usd": "$", "event_exchange_rate": "Event exchange rate",
        "event_exchange_help": "1 $ = how many FC?",
        "Performance globale": "Overall performance", "Bienvenue": "Welcome", "Nouvelle vente": "New sale",
        "Événements": "Events", "Ventes": "Sales", "Billets": "Tickets", "Revenu total": "Total revenue",
        "Événements à venir": "Upcoming events", "Prix moyen billet": "Average ticket price",
        "Taux de remplissage": "Occupancy rate", "Billets validés": "Checked-in tickets",
        "Programmation active": "Active schedule", "Par vente": "Per sale", "Capacité moyenne": "Average capacity",
        "Contrôle d’entrée": "Entry control", "Transactions": "Transactions", "Ventes récentes": "Recent sales",
        "Activité": "Activity", "Derniers billets générés": "Latest tickets generated", "Performance": "Performance",
        "Points clés": "Key insights", "Aucun événement pour le moment.": "No events yet.",
        "Liste des événements": "Event list", "Ajouter un événement": "Add event",
        "Catégories et prix": "Categories and prices", "Ajouter une catégorie": "Add category",
        "Vente de billets": "Ticket sales", "Sélectionner": "Select", "Valider la vente": "Confirm sale",
        "Billets vendus": "Sold tickets", "Contrôle d’entrée": "Entry control", "Valider l’entrée": "Validate entry",
        "Rapports de ventes": "Sales reports", "Rapport journalier": "Daily report", "Rapport mensuel": "Monthly report",
        "Filtrer": "Filter", "Télécharger PDF": "Download PDF", "Télécharger Excel": "Download Excel",
        "Chiffre d’affaires": "Revenue", "Nombre de ventes": "Number of sales", "Billets vendus": "Tickets sold",
    },
}


def current_language():
    return session.get("language", "fr") if session.get("language", "fr") in TRANSLATIONS else "fr"


def translate(key):
    return TRANSLATIONS[current_language()].get(key, key)


def current_currency():
    return session.get("currency", "FC") if session.get("currency", "FC") in ("FC", "USD") else "FC"


def convert_amount(amount, event=None):
    amount = float(amount or 0)
    if current_currency() == "USD":
        rate = float(getattr(event, "taux_conversion", 2800) or 2800) if event else 2800
        return amount / rate
    return amount


def money(amount, event=None):
    symbol = "$" if current_currency() == "USD" else "FC"
    return f"{convert_amount(amount, event):,.2f} {symbol}"


@app.context_processor
def inject_preferences():
    return {
        "_": translate,
        "money": money,
        "current_currency": current_currency(),
        "current_language": current_language(),
    }


class User(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(80), unique=True, nullable=False)
    password = db.Column(db.String(200), nullable=False)
    role = db.Column(db.String(50), default="admin")
    language = db.Column(db.String(5), default="fr")
    currency = db.Column(db.String(5), default="FC")
    photo = db.Column(db.Text, nullable=True)


class Evenement(db.Model):
    __tablename__ = "evenement"
    id = db.Column(db.Integer, primary_key=True)
    nom = db.Column(db.String(150), nullable=False)
    date = db.Column(db.String(20), nullable=False)
    lieu = db.Column(db.String(150), nullable=False)
    capacite = db.Column(db.Integer, nullable=False)
    description = db.Column(db.Text, default="")
    taux_conversion = db.Column(db.Float, nullable=False, default=2800)
    categories = db.relationship("Categorie", backref="evenement", cascade="all, delete-orphan")
    tickets = db.relationship("Ticket", backref="evenement", cascade="all, delete-orphan")
    ventes = db.relationship("Vente", backref="evenement", cascade="all, delete-orphan")


class Categorie(db.Model):
    __tablename__ = "categorie"
    id = db.Column(db.Integer, primary_key=True)
    evenement_id = db.Column(db.Integer, db.ForeignKey("evenement.id"), nullable=False)
    nom = db.Column(db.String(100), nullable=False)
    prix = db.Column(db.Float, nullable=False, default=0.0)
    quantite = db.Column(db.Integer, nullable=False, default=0)
    tickets = db.relationship("Ticket", backref="categorie", cascade="all, delete-orphan")


class Vente(db.Model):
    __tablename__ = "vente"
    id = db.Column(db.Integer, primary_key=True)
    evenement_id = db.Column(db.Integer, db.ForeignKey("evenement.id"), nullable=False)
    acheteur = db.Column(db.String(150), nullable=False)
    email = db.Column(db.String(150), nullable=False)
    total = db.Column(db.Float, nullable=False, default=0.0)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    tickets = db.relationship("Ticket", backref="vente", cascade="all, delete-orphan")


class Ticket(db.Model):
    __tablename__ = "ticket"
    id = db.Column(db.Integer, primary_key=True)
    reference = db.Column(db.String(30), unique=True, nullable=False)
    evenement_id = db.Column(db.Integer, db.ForeignKey("evenement.id"), nullable=False)
    categorie_id = db.Column(db.Integer, db.ForeignKey("categorie.id"), nullable=False)
    vente_id = db.Column(db.Integer, db.ForeignKey("vente.id"), nullable=False)
    acheteur = db.Column(db.String(150), nullable=False)
    email = db.Column(db.String(150), nullable=False)
    prix = db.Column(db.Float, nullable=False)
    checked_in = db.Column(db.Boolean, default=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)


def generate_reference():
    return "TCK-" + "".join(random.choices(string.ascii_uppercase + string.digits, k=8))


def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if "user_id" not in session:
            return redirect(url_for("login"))
        return f(*args, **kwargs)

    return decorated_function


def role_required(*allowed_roles):
    allowed_roles = tuple(role.lower() for role in allowed_roles)

    def decorator(f):
        @wraps(f)
        def decorated_function(*args, **kwargs):
            if "user_id" not in session:
                return redirect(url_for("login"))

            user = User.query.get(session["user_id"])
            if user is None:
                session.clear()
                return redirect(url_for("login"))

            role = (user.role or "").lower()
            if role == "admin":
                role = "administrateur"
                user.role = "administrateur"
                db.session.commit()

            if role not in allowed_roles:
                flash("Accès refusé pour votre rôle.", "danger")
                return redirect(url_for("dashboard"))
            session["user_role"] = role
            return f(*args, **kwargs)

        return decorated_function

    return decorator


@app.route("/")
def index():
    if "user_id" in session:
        return redirect(url_for("dashboard"))
    return render_template("index.html")


@app.route("/accueil")
def accueil():
    if "user_id" in session:
        return redirect(url_for("dashboard"))
    return render_template("index.html")


@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")
        user = User.query.filter_by(username=username, password=password).first()
        if user:
            if (user.role or "").lower() == "admin":
                user.role = "administrateur"
                db.session.commit()
            session["user_id"] = user.id
            session["username"] = user.username
            session["user_role"] = (user.role or "administrateur").lower()
            session["language"] = user.language or "fr"
            session["currency"] = user.currency or "FC"
            flash(f"Bienvenue, {user.username} !", "success")
            return redirect(url_for("dashboard"))
        flash("Identifiants incorrects.", "danger")
    return render_template("login.html")


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


@app.route("/account", methods=["GET", "POST"])
@login_required
def account():
    user = User.query.get(session["user_id"])
    if request.method == "POST":
        user.language = request.form.get("language", "fr") if request.form.get("language") in TRANSLATIONS else "fr"
        user.currency = request.form.get("currency", "FC") if request.form.get("currency") in ("FC", "USD") else "FC"
        db.session.commit()
        session["language"] = user.language
        session["currency"] = user.currency
        flash(translate("settings_saved"), "success")
        return redirect(url_for("account"))
    return render_template("account.html", user=user)


@app.route("/a-propos")
def a_propos():
    return render_template("a_propos.html")


@app.route("/dashboard")
@login_required
@role_required("administrateur", "auditeur", "vendeur")
def dashboard():
    total_evenements = Evenement.query.count()
    total_ventes = Vente.query.count()
    total_billets = Ticket.query.count()
    billets_valides = Ticket.query.filter_by(checked_in=True).count()
    total_revenu = db.session.query(func.coalesce(func.sum(Vente.total), 0)).scalar() or 0
    ventes_recents = Vente.query.order_by(Vente.created_at.desc()).limit(5).all()
    avg_ticket_price = db.session.query(func.avg(Ticket.prix)).scalar() or 0
    upcoming_events = 0
    today = datetime.utcnow().strftime("%Y-%m-%d")
    for evenement in Evenement.query.all():
        if evenement.date and evenement.date >= today:
            upcoming_events += 1
    occupancy = 0
    if total_evenements:
        occupancy = round((sum((Ticket.query.filter_by(evenement_id=e.id).count() / (e.capacite or 1)) for e in Evenement.query.all()) / total_evenements) * 100, 1)
    checked_rate = round((billets_valides / total_billets) * 100, 1) if total_billets else 0

    recent_activity = Ticket.query.order_by(Ticket.created_at.desc()).limit(4).all()

    event_sales = (
        db.session.query(Evenement.nom, func.count(Ticket.id).label("tickets_vendus"))
        .outerjoin(Ticket, Ticket.evenement_id == Evenement.id)
        .group_by(Evenement.id, Evenement.nom)
        .order_by(func.count(Ticket.id).desc())
        .all()
    )
    event_labels = [item[0] for item in event_sales]
    event_values = [item[1] for item in event_sales]

    status_billets = {
        "Validés": billets_valides,
        "En attente": max(total_billets - billets_valides, 0),
    }

    category_revenue = (
        db.session.query(Categorie.nom, func.coalesce(func.sum(Ticket.prix), 0).label("revenu"))
        .outerjoin(Ticket, Ticket.categorie_id == Categorie.id)
        .group_by(Categorie.id, Categorie.nom)
        .order_by(func.coalesce(func.sum(Ticket.prix), 0).desc())
        .all()
    )
    category_labels = [item[0] for item in category_revenue]
    category_values = [float(item[1]) for item in category_revenue]

    event_load = []
    for evenement in Evenement.query.order_by(Evenement.date).all():
        vendus = Ticket.query.filter_by(evenement_id=evenement.id).count()
        capacite = evenement.capacite or 1
        taux = min(round((vendus / capacite) * 100), 100) if capacite else 0
        event_load.append({
            "nom": evenement.nom,
            "vendus": vendus,
            "capacite": evenement.capacite,
            "taux": taux,
        })

    return render_template(
        "dashboard.html",
        total_evenements=total_evenements,
        total_ventes=total_ventes,
        total_billets=total_billets,
        billets_valides=billets_valides,
        total_revenu=total_revenu,
        ventes_recents=ventes_recents,
        event_labels=event_labels,
        event_values=event_values,
        status_billets=status_billets,
        category_labels=category_labels,
        category_values=category_values,
        event_load=event_load,
        avg_ticket_price=avg_ticket_price,
        upcoming_events=upcoming_events,
        occupancy=occupancy,
        checked_rate=checked_rate,
        recent_activity=recent_activity,
    )


def report_period():
    period = request.args.get("period", "month")
    if period not in ("day", "month"):
        period = "month"
    now = datetime.utcnow()
    if period == "day":
        selected = request.args.get("date", now.strftime("%Y-%m-%d"))
        try:
            start = datetime.strptime(selected, "%Y-%m-%d")
        except ValueError:
            selected = now.strftime("%Y-%m-%d")
            start = datetime.strptime(selected, "%Y-%m-%d")
        end = start.replace(hour=23, minute=59, second=59)
        label = selected
    else:
        selected = request.args.get("month", now.strftime("%Y-%m"))
        try:
            start = datetime.strptime(selected, "%Y-%m")
        except ValueError:
            selected = now.strftime("%Y-%m")
            start = datetime.strptime(selected, "%Y-%m")
        end = datetime(start.year + (start.month == 12), 1 if start.month == 12 else start.month + 1, 1)
        label = selected
    return period, selected, start, end, label


def report_data():
    period, selected, start, end, label = report_period()
    sales = Vente.query.filter(Vente.created_at >= start, Vente.created_at < end).order_by(Vente.created_at.desc()).all()
    total_revenue = sum(float(sale.total or 0) for sale in sales)
    total_tickets = sum(len(sale.tickets) for sale in sales)
    return {
        "period": period,
        "selected": selected,
        "label": label,
        "sales": sales,
        "total_revenue": total_revenue,
        "total_tickets": total_tickets,
    }


@app.route("/rapports")
@login_required
@role_required("administrateur", "auditeur")
def rapports():
    return render_template("rapports.html", report=report_data())


@app.route("/rapports/export/excel")
@login_required
@role_required("administrateur", "auditeur")
def export_excel():
    report = report_data()
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Rapport ventes"
    sheet.append(["StadeControl - Stade Communal"])
    sheet.append(["Rapport", report["label"]])
    sheet.append([])
    sheet.append(["Date", "Client", "Email", "Evenement", "Billets", "Total"])
    for sale in report["sales"]:
        sheet.append([
            sale.created_at.strftime("%Y-%m-%d %H:%M"), sale.acheteur, sale.email,
            sale.evenement.nom, len(sale.tickets), sale.total,
        ])
    sheet.append([])
    sheet.append(["Totaux", "", "", "", report["total_tickets"], report["total_revenue"]])
    sheet["A1"].font = Font(bold=True, size=16, color="FFFFFF")
    for cell in sheet["1:1"]:
        cell.fill = PatternFill("solid", fgColor="1D1B2D")
    for cell in sheet["4:4"]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="8B5E3C")
    for column in sheet.columns:
        sheet.column_dimensions[column[0].column_letter].width = min(max(len(str(cell.value or "")) for cell in column) + 2, 34)
    output = BytesIO()
    workbook.save(output)
    output.seek(0)
    filename = f"rapport-{report['period']}-{report['selected']}.xlsx"
    return send_file(output, as_attachment=True, download_name=filename, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.route("/rapports/export/pdf")
@login_required
@role_required("administrateur", "auditeur")
def export_pdf():
    report = report_data()
    output = BytesIO()
    document = SimpleDocTemplate(output, pagesize=landscape(A4), rightMargin=28, leftMargin=28, topMargin=28, bottomMargin=28)
    styles = getSampleStyleSheet()
    content = [Paragraph("StadeControl - Stade Communal", styles["Title"]), Paragraph(f"Rapport des ventes : {report['label']}", styles["Heading2"]), Spacer(1, 14)]
    rows = [["Date", "Client", "Email", "Evenement", "Billets", "Total"]]
    for sale in report["sales"]:
        rows.append([
            sale.created_at.strftime("%Y-%m-%d %H:%M"), sale.acheteur, sale.email,
            sale.evenement.nom, str(len(sale.tickets)), money(sale.total, sale.evenement),
        ])
    rows.append(["Totaux", "", "", "", str(report["total_tickets"]), money(report["total_revenue"])])
    table = Table(rows, repeatRows=1)
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1D1B2D")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("BACKGROUND", (0, -1), (-1, -1), colors.HexColor("#F1E7DF")),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#D9CEC5")),
        ("PADDING", (0, 0), (-1, -1), 7),
    ]))
    content.append(table)
    document.build(content)
    output.seek(0)
    filename = f"rapport-{report['period']}-{report['selected']}.pdf"
    return send_file(output, as_attachment=True, download_name=filename, mimetype="application/pdf")


@app.route("/evenements")
@login_required
@role_required("administrateur", "auditeur")
def evenements():
    liste_evenements = Evenement.query.order_by(Evenement.date).all()
    return render_template("evenements.html", evenements=liste_evenements)


@app.route("/evenements/ajouter", methods=["GET", "POST"])
@login_required
@role_required("administrateur")
def ajouter_evenement():
    if request.method == "POST":
        nouveau_evenement = Evenement(
            nom=request.form["nom"],
            date=request.form["date"],
            lieu=request.form["lieu"],
            capacite=int(request.form["capacite"]),
            description=request.form.get("description", ""),
            taux_conversion=float(request.form.get("taux_conversion", 2800) or 2800),
        )
        db.session.add(nouveau_evenement)
        db.session.commit()
        flash("Événement ajouté avec succès.", "success")
        return redirect(url_for("evenements"))
    return render_template("nouvel_evenement.html")


@app.route("/categories")
@login_required
@role_required("administrateur")
def categories():
    evenements = Evenement.query.order_by(Evenement.date).all()
    selected_event_id = request.args.get("evenement_id", type=int)
    if selected_event_id:
        categories = Categorie.query.filter_by(evenement_id=selected_event_id).order_by(Categorie.nom).all()
    else:
        categories = Categorie.query.order_by(Categorie.nom).all()
    return render_template(
        "categories.html",
        evenements=evenements,
        categories=categories,
        selected_event_id=selected_event_id,
    )


@app.route("/categories/ajouter", methods=["GET", "POST"])
@login_required
@role_required("administrateur")
def ajouter_categorie():
    if request.method == "POST":
        categorie = Categorie(
            evenement_id=int(request.form["evenement_id"]),
            nom=request.form["nom"],
            prix=float(request.form["prix"]),
            quantite=int(request.form["quantite"]),
        )
        db.session.add(categorie)
        db.session.commit()
        flash("Catégorie ajoutée.", "success")
        return redirect(url_for("categories", evenement_id=categorie.evenement_id))
    evenements = Evenement.query.order_by(Evenement.date).all()
    return render_template("nouvel_categorie.html", evenements=evenements)


@app.route("/ventes", methods=["GET", "POST"])
@login_required
@role_required("administrateur", "vendeur")
def ventes():
    evenements = Evenement.query.order_by(Evenement.date).all()
    categories = []
    selected_event_id = request.args.get("evenement_id", type=int)

    if request.method == "POST":
        evenement_id = int(request.form["evenement_id"])
        categorie_id = int(request.form["categorie_id"])
        categorie = Categorie.query.get_or_404(categorie_id)
        quantite = int(request.form["quantite"])
        acheteur = request.form["acheteur"]
        email = request.form["email"]

        vente = Vente(
            evenement_id=evenement_id,
            acheteur=acheteur,
            email=email,
            total=categorie.prix * quantite,
        )
        db.session.add(vente)
        db.session.flush()

        for _ in range(quantite):
            ticket = Ticket(
                reference=generate_reference(),
                evenement_id=evenement_id,
                categorie_id=categorie_id,
                vente_id=vente.id,
                acheteur=acheteur,
                email=email,
                prix=categorie.prix,
            )
            db.session.add(ticket)

        db.session.commit()
        flash(f"Vente enregistrée : {quantite} billet(s) créés.", "success")
        return redirect(url_for("billets"))

    if selected_event_id:
        categories = Categorie.query.filter_by(evenement_id=selected_event_id).all()

    return render_template(
        "ventes.html",
        evenements=evenements,
        categories=categories,
        selected_event_id=selected_event_id,
    )


@app.route("/billets")
@login_required
@role_required("administrateur", "auditeur", "vendeur")
def billets():
    tickets = Ticket.query.order_by(Ticket.created_at.desc()).all()
    return render_template("billets.html", tickets=tickets)


@app.route("/qr/<int:ticket_id>")
@login_required
@role_required("administrateur", "auditeur", "vendeur")
def qr_code(ticket_id):
    ticket = Ticket.query.get_or_404(ticket_id)
    qr_url = f"https://api.qrserver.com/v1/create-qr-code/?size=220x220&data={ticket.reference}"
    return render_template("qr_code.html", ticket=ticket, qr_url=qr_url)


@app.route("/controle-entree", methods=["GET", "POST"])
@login_required
@role_required("administrateur", "auditeur")
def controle_entree():
    ticket = None
    message = None
    status = "info"

    if request.method == "POST":
        reference = request.form.get("reference", "").strip()
        ticket = Ticket.query.filter_by(reference=reference).first()

        if ticket is None:
            message = "Billet introuvable. Vérifiez la référence."
            status = "danger"
        elif ticket.checked_in:
            message = "Ce billet a déjà été validé à l’entrée."
            status = "warning"
        else:
            ticket.checked_in = True
            db.session.commit()
            message = f"Entrée validée pour {ticket.acheteur}."
            status = "success"

    return render_template("controle_entree.html", ticket=ticket, message=message, status=status)


@app.route("/users", methods=["GET", "POST"])
@login_required
@role_required("administrateur")
def users():
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")
        role = request.form.get("role", "vendeur")
        photo = request.form.get("photo", "")
        if photo and not photo.startswith("data:image/"):
            photo = ""

        if username and password:
            if User.query.filter_by(username=username).first():
                flash("Ce nom d’utilisateur existe déjà.", "danger")
            else:
                user = User(username=username, password=password, role=role, photo=photo or None)
                db.session.add(user)
                db.session.commit()
                flash(f"Utilisateur {username} créé avec le rôle {role}.", "success")
        else:
            flash("Merci de remplir le nom et le mot de passe.", "danger")

    users_list = User.query.order_by(User.username).all()
    return render_template("users.html", users=users_list)


with app.app_context():
    db.create_all()

    evenement_columns = [col[1] for col in db.session.execute(text("PRAGMA table_info(evenement)"))]
    if "description" not in evenement_columns:
        db.session.execute(text("ALTER TABLE evenement ADD COLUMN description TEXT DEFAULT ''"))
        db.session.commit()
    if "taux_conversion" not in evenement_columns:
        db.session.execute(text("ALTER TABLE evenement ADD COLUMN taux_conversion FLOAT DEFAULT 2800"))
        db.session.commit()

    user_columns = [col[1] for col in db.session.execute(text("PRAGMA table_info(user)"))]
    if "role" not in user_columns:
        db.session.execute(text("ALTER TABLE user ADD COLUMN role VARCHAR(50) DEFAULT 'administrateur'"))
        db.session.commit()
    if "language" not in user_columns:
        db.session.execute(text("ALTER TABLE user ADD COLUMN language VARCHAR(5) DEFAULT 'fr'"))
        db.session.commit()
    if "currency" not in user_columns:
        db.session.execute(text("ALTER TABLE user ADD COLUMN currency VARCHAR(5) DEFAULT 'FC'"))
        db.session.commit()
    if "photo" not in user_columns:
        db.session.execute(text("ALTER TABLE user ADD COLUMN photo TEXT"))
        db.session.commit()

    for user in User.query.all():
        if (user.role or "").lower() == "admin":
            user.role = "administrateur"
        if not user.language:
            user.language = "fr"
        if not user.currency:
            user.currency = "FC"
    for evenement in Evenement.query.all():
        if not evenement.taux_conversion:
            evenement.taux_conversion = 2800
    db.session.commit()

    if not User.query.filter_by(username="admin").first():
        db.session.add(User(username="admin", password="admin123", role="administrateur"))
        db.session.commit()


if __name__ == "__main__":
    app.run(debug=True)
