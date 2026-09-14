from flask import Flask, redirect, render_template, request, url_for
from flask_sqlalchemy import SQLAlchemy


app = Flask(__name__)
app.config["SQLALCHEMY_DATABASE_URI"] = "sqlite:///database.db"
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False
db = SQLAlchemy(app)


class Evenement(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    nom = db.Column(db.String(150), nullable=False)
    date = db.Column(db.String(20), nullable=False)
    lieu = db.Column(db.String(150), nullable=False)
    capacite = db.Column(db.Integer, nullable=False)


@app.route("/")
def accueil():
    return render_template("index.html")


@app.route("/evenements")
def evenements():
    liste_evenements = Evenement.query.order_by(Evenement.date).all()
    return render_template("evenements.html", evenements=liste_evenements)


@app.route("/evenements/ajouter", methods=["GET", "POST"])
def ajouter_evenement():
    if request.method == "POST":
        nouvel_evenement = Evenement(
            nom=request.form["nom"],
            date=request.form["date"],
            lieu=request.form["lieu"],
            capacite=int(request.form["capacite"]),
        )
        db.session.add(nouvel_evenement)
        db.session.commit()
        return redirect(url_for("evenements"))

    return render_template("nouvel_evenement.html")


with app.app_context():
    db.create_all()


if __name__ == "__main__":
    app.run(debug=True)
