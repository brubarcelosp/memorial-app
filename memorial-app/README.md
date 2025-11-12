### Step 1: Set Up Your Environment

1. **Install Flask**: Make sure you have Flask installed. You can install it using pip:

   ```bash
   pip install Flask Flask-SQLAlchemy Flask-Login
   ```

2. **Create the Project Structure**:

   ```
   flask_app/
   ├── app.py
   ├── models.py
   ├── forms.py
   ├── templates/
   │   ├── base.html
   │   ├── login.html
   │   ├── register.html
   │   └── dashboard.html
   └── static/
       └── style.css
   ```

### Step 2: Create the Application

#### `app.py`

```python
from flask import Flask, render_template, redirect, url_for, flash, request
from flask_sqlalchemy import SQLAlchemy
from flask_login import LoginManager, UserMixin, login_user, login_required, logout_user, current_user
from forms import RegistrationForm, LoginForm

app = Flask(__name__)
app.config['SECRET_KEY'] = 'your_secret_key'
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///site.db'
db = SQLAlchemy(app)
login_manager = LoginManager(app)
login_manager.login_view = 'login'

class User(db.Model, UserMixin):
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(150), unique=True, nullable=False)
    password = db.Column(db.String(150), nullable=False)

@login_manager.user_loader
def load_user(user_id):
    return User.query.get(int(user_id))

@app.route('/')
def home():
    return render_template('base.html')

@app.route('/register', methods=['GET', 'POST'])
def register():
    form = RegistrationForm()
    if form.validate_on_submit():
        user = User(username=form.username.data, password=form.password.data)
        db.session.add(user)
        db.session.commit()
        flash('Account created successfully!', 'success')
        return redirect(url_for('login'))
    return render_template('register.html', form=form)

@app.route('/login', methods=['GET', 'POST'])
def login():
    form = LoginForm()
    if form.validate_on_submit():
        user = User.query.filter_by(username=form.username.data).first()
        if user and user.password == form.password.data:  # In production, use hashed passwords
            login_user(user)
            return redirect(url_for('dashboard'))
        else:
            flash('Login Unsuccessful. Please check username and password', 'danger')
    return render_template('login.html', form=form)

@app.route('/dashboard')
@login_required
def dashboard():
    return render_template('dashboard.html', username=current_user.username)

@app.route('/logout')
@login_required
def logout():
    logout_user()
    return redirect(url_for('home'))

if __name__ == '__main__':
    db.create_all()  # Create database tables
    app.run(debug=True)
```

#### `models.py`

This file is already included in `app.py` as the `User` model.

#### `forms.py`

```python
from flask_wtf import FlaskForm
from wtforms import StringField, PasswordField, SubmitField
from wtforms.validators import DataRequired, Length, EqualTo

class RegistrationForm(FlaskForm):
    username = StringField('Username', validators=[DataRequired(), Length(min=2, max=150)])
    password = PasswordField('Password', validators=[DataRequired(), Length(min=6, max=150)])
    submit = SubmitField('Sign Up')

class LoginForm(FlaskForm):
    username = StringField('Username', validators=[DataRequired()])
    password = PasswordField('Password', validators=[DataRequired()])
    submit = SubmitField('Login')
```

#### HTML Templates

1. **`base.html`**:

```html
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <link rel="stylesheet" href="{{ url_for('static', filename='style.css') }}">
    <title>Flask App</title>
</head>
<body>
    <nav>
        <a href="{{ url_for('home') }}">Home</a>
        {% if current_user.is_authenticated %}
            <a href="{{ url_for('dashboard') }}">Dashboard</a>
            <a href="{{ url_for('logout') }}">Logout</a>
        {% else %}
            <a href="{{ url_for('login') }}">Login</a>
            <a href="{{ url_for('register') }}">Register</a>
        {% endif %}
    </nav>
    <div class="container">
        {% with messages = get_flashed_messages() %}
            {% if messages %}
                <ul>
                {% for message in messages %}
                    <li>{{ message }}</li>
                {% endfor %}
                </ul>
            {% endif %}
        {% endwith %}
        {% block content %}{% endblock %}
    </div>
</body>
</html>
```

2. **`login.html`**:

```html
{% extends 'base.html' %}
{% block content %}
<h2>Login</h2>
<form method="POST">
    {{ form.hidden_tag() }}
    {{ form.username.label }} {{ form.username(size=32) }}<br>
    {{ form.password.label }} {{ form.password(size=32) }}<br>
    {{ form.submit() }}
</form>
{% endblock %}
```

3. **`register.html`**:

```html
{% extends 'base.html' %}
{% block content %}
<h2>Register</h2>
<form method="POST">
    {{ form.hidden_tag() }}
    {{ form.username.label }} {{ form.username(size=32) }}<br>
    {{ form.password.label }} {{ form.password(size=32) }}<br>
    {{ form.submit() }}
</form>
{% endblock %}
```

4. **`dashboard.html`**:

```html
{% extends 'base.html' %}
{% block content %}
<h2>Welcome, {{ username }}!</h2>
<p>This is your dashboard.</p>
{% endblock %}
```

#### `static/style.css`

You can add some basic styles here.

### Step 3: Run the Application

1. Navigate to your project directory.
2. Run the application:

   ```bash
   python app.py
   ```

3. Open your web browser and go to `http://127.0.0.1:5000/`.

### Conclusion

This is a basic Flask web application that includes user registration, login, and a dashboard. You can expand upon this by adding features such as password hashing, email verification, and more. Make sure to handle user passwords securely in a production environment by using libraries like `bcrypt` or `werkzeug.security` for hashing.