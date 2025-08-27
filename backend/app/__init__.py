import os
from dotenv import load_dotenv

# Get the directory where this script is located
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Build absolute path to the .env.dev file
ENV_PATH = os.path.join(os.path.dirname(BASE_DIR), ".env.dev")

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ENV_PATH = os.path.join(BASE_DIR, ".env.dev")

load_dotenv(ENV_PATH, override=True)
