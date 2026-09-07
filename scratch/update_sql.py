import json

with open('src/data/planesData.js', 'r', encoding='utf-8') as f:
    code = f.read()

json_str = code.split('export const PLANES_MANTENIMIENTO_DATA = ')[1].strip().rstrip(';')
data = json.loads(json_str)

with open('supabase/setup_compact.sql', 'r', encoding='utf-8') as f:
    sql = f.read()

costos_rows = []
for p in data:
    for km_str, c in sorted(p['costosPorKm'].items(), key=lambda x: int(x[0])):
        costos_rows.append(f"((SELECT id FROM public.planes_mantenimiento WHERE codigo_plan = '{p['codigo_plan']}'), {c['km']}, {c['total_repuestos']}, {c['total_lubricantes']}, {c['total_mano_obra']}, {c['total_km']})")

print(f"Total plan_costos_km rows: {len(costos_rows)}")

header = 'INSERT INTO public.plan_costos_km (plan_id, km, total_repuestos, total_lubricantes, total_mano_obra, total_km) VALUES\n'
footer = '\nON CONFLICT (plan_id, km) DO UPDATE SET total_repuestos = EXCLUDED.total_repuestos, total_lubricantes = EXCLUDED.total_lubricantes, total_mano_obra = EXCLUDED.total_mano_obra, total_km = EXCLUDED.total_km;'

parts = sql.split('INSERT INTO public.plan_costos_km (plan_id, km, total_repuestos, total_lubricantes, total_mano_obra, total_km) VALUES')
after_parts = parts[1].split('ON CONFLICT (plan_id, km) DO UPDATE SET total_repuestos = EXCLUDED.total_repuestos')

rest_of_sql = after_parts[1].split(';', 1)[1]

new_sql = parts[0] + header + ',\n'.join(costos_rows) + footer + rest_of_sql

with open('supabase/setup_compact.sql', 'w', encoding='utf-8') as f:
    f.write(new_sql)

print('✅ supabase/setup_compact.sql updated successfully!')
