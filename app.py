from flask import Flask
import logging
from upload import upload_blueprint
from generate import generate_blueprint

# Set up logging
logging.basicConfig(level=logging.DEBUG)

app = Flask(__name__)

# Register Blueprints
app.register_blueprint(upload_blueprint, url_prefix="/")
app.register_blueprint(generate_blueprint, url_prefix="/generate_docs")

if __name__ == "__main__":
    app.run(debug=True)
