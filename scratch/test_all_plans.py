import json

with open('src/data/planesData.js', 'r', encoding='utf-8') as f:
    code = f.read()

json_str = code.split('export const PLANES_MANTENIMIENTO_DATA = ')[1].strip().rstrip(';')
data = json.loads(json_str)

for p in data:
    kms = sorted([int(k) for k in p['costosPorKm'].keys()])
    c5k = p['costosPorKm'].get('5000') or {}
    c60k = p['costosPorKm'].get('60000') or {}
    print(f"Plan: {p['id']:<20} | {p['nombre_completo']:<35} | KMs: {len(kms)} (from {kms[0]} to {kms[-1]})")
    print(f"   5k: rep={c5k.get('total_repuestos')} lub={c5k.get('total_lubricantes')} mo={c5k.get('total_mano_obra')} tot={c5k.get('total_km')}")
    print(f"  60k: rep={c60k.get('total_repuestos')} lub={c60k.get('total_lubricantes')} mo={c60k.get('total_mano_obra')} tot={c60k.get('total_km')}")

print("\n--- TEST TANK 300 (3 VEHICULOS, 5000 a 60000 KM) ---")
tank = next(p for p in data if p['id'] == 'tank-300-120k')
rep_1v = sum(c['total_repuestos'] for k, c in tank['costosPorKm'].items() if 5000 <= int(k) <= 60000)
lub_1v = sum(c['total_lubricantes'] for k, c in tank['costosPorKm'].items() if 5000 <= int(k) <= 60000)
mo_1v  = sum(c['total_mano_obra'] for k, c in tank['costosPorKm'].items() if 5000 <= int(k) <= 60000)
prev_1v = rep_1v + lub_1v + mo_1v

rep_3v = round(rep_1v * 3, 2)
lub_3v = round(lub_1v * 3, 2)
mo_3v  = round(mo_1v * 3, 2)
prev_3v = round(prev_1v * 3, 2)
corr_3v = round(prev_3v * 0.30, 2)
gran_total = round(prev_3v + corr_3v, 2)

print(f"Total Repuestos 3V:    ${rep_3v:,.2f}  (Esperado: $4,746.90)")
print(f"Total Lubricantes 3V:  ${lub_3v:,.2f}  (Esperado: $4,282.35)")
print(f"Total Mano de Obra 3V: ${mo_3v:,.2f}  (Esperado: $7,828.20)")
print(f"Total Preventivo 3V:   ${prev_3v:,.2f} (Esperado: $16,857.45)")
print(f"Correctivo (30%):      ${corr_3v:,.2f}  (Esperado: $5,057.24)")
print(f"Gran Total (1+2+3+4):  ${gran_total:,.2f} (Esperado: $21,914.69)")
