from datetime import datetime
from functools import wraps
import random
import string

from flask import Flask, flash, redirect, render_template, request, session, url_for
from flask_sqlalchemy import SQLAlchemy
from sqlalchemy import text


app = Flask(__name__)
app.config["SECRET_KEY"] = "ticket-secret-key"
app.config["SQLALCHEMY_DATABASE_URI"] = "sqlite:///database.db"
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False
db = SQLAlchemy(app)


class User(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(80), unique=True, nullable=False)
    password = db.Column(db.String(200), nullable=False)
    role = db.Column(db.String(50), default="admin")


class Evenement(db.Model):
    __tablename__ = "evenement"
    id = db.Column(db.Integer, primary_key=True)
    nom = db.Column(db.String(150), nullable=False)
    date = db.Column(db.String(20), nullable=False)
    lieu = db.Column(db.String(150), nullable=False)
    capacite = db.Column(db.Integer, nullable=False)
    description = db.Column(db.Text, default="")
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


@app.route("/")
def index():
    if "user_id" in session:
        return redirect(url_for("dashboard"))
    return redirect(url_for("login"))


@app.route("/accueil")
def accueil():
    return redirect(url_for("dashboard")) if "user_id" in session else redirect(url_for("login"))


@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")
        user = User.query.filter_by(username=username, password=password).first()
        if user:
            session["user_id"] = user.id
            session["username"] = user.username
            flash("Connexion réussie.", "success")
            return redirect(url_for("dashboard"))
        flash("Identifiants incorrects.", "danger")
    return render_template("login.html")


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


@app.route("/dashboard")
@login_required
def dashboard():
    total_evenements = Evenement.query.count()
    total_ventes = Vente.query.count()
    total_billets = Ticket.query.count()
    billets_valides = Ticket.query.filter_by(checked_in=True).count()
    ventes_recents = Vente.query.order_by(Vente.created_at.desc()).limit(5).all()
    return render_template(
        "dashboard.html",
        total_evenements=total_evenements,
        total_ventes=total_ventes,
        total_billets=total_billets,
        billets_valides=billets_valides,
        ventes_recents=ventes_recents,
    )


@app.route("/evenements")
@login_required
def evenements():
    liste_evenements = Evenement.query.order_by(Evenement.date).all()
    return render_template("evenements.html", evenements=liste_evenements)


@app.route("/evenements/ajouter", methods=["GET", "POST"])
@login_required
def ajouter_evenement():
    if request.method == "POST":
        nouveau_evenement = Evenement(
            nom=request.form["nom"],
            date=request.form["date"],
            lieu=request.form["lieu"],
            capacite=int(request.form["capacite"]),
            description=request.form.get("description", ""),
        )
        db.session.add(nouveau_evenement)
        db.session.commit()
        flash("Événement ajouté avec succès.", "success")
        return redirect(url_for("evenements"))
    return render_template("nouvel_evenement.html")


@app.route("/categories")
@login_required
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
def billets():
    tickets = Ticket.query.order_by(Ticket.created_at.desc()).all()
    return render_template("billets.html", tickets=tickets)


@app.route("/qr/<int:ticket_id>")
@login_required
def qr_code(ticket_id):
    ticket = Ticket.query.get_or_404(ticket_id)
    qr_url = f"https://api.qrserver.com/v1/create-qr-code/?size=220x220&data={ticket.reference}"
    return render_template("qr_code.html", ticket=ticket, qr_url=qr_url)


@app.route("/controle-entree", methods=["GET", "POST"])
@login_required
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


with app.app_context():
    db.create_all()

    evenement_columns = [col[1] for col in db.session.execute(text("PRAGMA table_info(evenement)"))]
    if "description" not in evenement_columns:
        db.session.execute(text("ALTER TABLE evenement ADD COLUMN description TEXT DEFAULT ''"))
        db.session.commit()

    if not User.query.filter_by(username="admin").first():
        db.session.add(User(username="admin", password="admin123", role="admin"))
        db.session.commit()


if __name__ == "__main__":
    app.run(debug=True)
