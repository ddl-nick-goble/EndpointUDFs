import os

def print_domino_envs():
    domino_envs = {
        k: v for k, v in os.environ.items()
        if k.startswith("DOMINO_")
    }

    if not domino_envs:
        print("No DOMINO_* environment variables found.")
        return

    for key in sorted(domino_envs):
        print(f"{key}={domino_envs[key]}")

print_domino_envs()
