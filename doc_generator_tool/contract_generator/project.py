import json
def save_project(p, path):
    with open(path, 'w') as f: json.dump(p.to_dict(), f)
def load_project(path):
    from .models import ProjectData
    with open(path, 'r') as f: return ProjectData(**json.load(f))
