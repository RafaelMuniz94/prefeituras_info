import os
from dotenv import load_dotenv

load_dotenv()

class Configuration:
    def __init__(self,):
        self.base_url = os.environ.get("BaseURL")