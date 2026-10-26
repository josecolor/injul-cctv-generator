import threading
import time
import requests


def iniciar_anti_sleep():
  """Función para mantener el servidor de Railway despierto

  haciendo un auto-ping cada 10 minutos (600 segundos).
  """

  def mantener_despierto():
    # URL pública de tu app en Railway
    url_app = "https://injul-cctv-generator-production.up.railway.app"
    while True:
      try:
        time.sleep(600)  # Espera 10 minutos
        requests.get(url_app)
      except Exception as e:
        print(f"Error en ping anti-sleep: {e}")

  # Arranca el proceso en un hilo secundario invisible y seguro
  hilo = threading.Thread(target=mantener_despierto, daemon=True)
  hilo.start()
