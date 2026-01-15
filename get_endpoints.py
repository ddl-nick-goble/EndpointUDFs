import requests
import os

DOMINO_URL = "https://se-demo.domino.tech"
API_KEY = os.environ.get("DOMINO_USER_API_KEY")

def get_project_models(project_id: str) -> list[dict]:
    """Fetch all models for a project."""
    url = f"{DOMINO_URL}/v4/modelManager/getModels"
    params = {"projectId": project_id}
    headers = {"X-Domino-Api-Key": API_KEY}
    
    resp = requests.get(url, params=params, headers=headers)
    resp.raise_for_status()
    return resp.json()

if __name__ == "__main__":
    import os

    project_id = os.environ.get("DOMINO_PROJECT_ID")
    print(project_id)

    models = get_project_models("69683eed65e9f57bdafc4f2e")
    
    for m in models:
        print(m)
        # print(f"{m['name']}: {m['activeVersionStatus']}")
        # print(f"  └─ Serving: {m['activeVersion']['registeredModelName']} v{m['activeVersion']['registeredModelVersion']}")
