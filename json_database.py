import json
import hashlib
import os

DB_FILE = "database.json"

print("json_database.py loaded")

def load_database():
    """Load the JSON database."""
    if not os.path.exists(DB_FILE):
        with open(DB_FILE, "w") as db_file:
            json.dump({"users": []}, db_file)
    with open(DB_FILE, "r") as db_file:
        return json.load(db_file)

def save_database(data):
    """Save data to the JSON database."""
    with open(DB_FILE, "w") as db_file:
        json.dump(data, db_file, indent=4)

def add_user(username, password, role="non-admin"):
    """Add a new user to the database."""
    db = load_database()
    hashed_password = hashlib.md5(password.encode()).hexdigest()
    if any(user["username"] == username for user in db["users"]):
        raise ValueError("Username already exists.")
    db["users"].append({"username": username, "password": hashed_password, "role": role})
    save_database(db)

def validate_user(username, password):
    hashed_password = hashlib.md5(password.encode()).hexdigest()
    with open("database.json", "r", encoding="utf-8") as f:
        data = json.load(f)
    for user in data.get("users", []):
        if user["username"] == username and user["password"] == hashed_password:
            return user.get("role", "non-admin")
    return None

def reset_admin_password(new_password):
    """Reset the admin password."""
    db = load_database()
    hashed_password = hashlib.md5(new_password.encode()).hexdigest()
    for user in db["users"]:
        if user["username"] == "admin":
            user["password"] = hashed_password
            break
    save_database(db)