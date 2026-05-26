import json, importlib, sys
sys.path.insert(0, "/tmp/opencode/condo_pet")

# Read and execute just the data part
with open("/tmp/opencode/condo_pet/generate_excel.py", "r") as f:
    content = f.read()

# Isolate the condos list by finding start/end
lines = content.split("\n")
in_condos = False
condo_lines = []
depth = 0
for line in lines:
    stripped = line.strip()
    if stripped.startswith("condos = ["):
        in_condos = True
        condo_lines.append(line)
        depth += stripped.count("[") - stripped.count("]")
        continue
    if in_condos:
        condo_lines.append(line)
        depth += stripped.count("[") - stripped.count("]")
        if depth <= 0 and stripped.startswith("]"):
            break

code = "\n".join(condo_lines)
exec(code)

output = []
for c in condos:
    output.append({
        "name": c.get("name", ""),
        "name_th": c.get("name_th", ""),
        "district": c.get("district", ""),
        "address": c.get("address", ""),
        "lat": c.get("lat"),
        "lon": c.get("lon"),
        "units": c.get("units"),
        "room_sizes": c.get("room_sizes", ""),
        "sale_price": c.get("sale_price", ""),
        "rent_price": c.get("rent_price", ""),
        "developer": c.get("developer", ""),
        "phone": c.get("phone", ""),
        "email": c.get("email", ""),
        "bts": c.get("bts", ""),
        "pet_policy": c.get("pet_policy", ""),
    })

print(json.dumps(output, ensure_ascii=False))
