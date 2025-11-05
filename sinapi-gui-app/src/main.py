# filepath: sinapi-gui-app/src/main.py

import sys
from ui import run_app

def main():
    # garante execução pelo ponto de entrada único
    run_app()

if __name__ == "__main__":
    main()