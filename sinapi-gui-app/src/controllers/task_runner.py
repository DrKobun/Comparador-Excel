import threading


def run_background(fn, args=(), kwargs=None):
    """Executa `fn` em thread daemon. Retorna o objeto Thread."""
    if kwargs is None:
        kwargs = {}
    t = threading.Thread(target=fn, args=args, kwargs=kwargs, daemon=True)
    t.start()
    return t
