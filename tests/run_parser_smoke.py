from parser_engine import parse_fat_inline, parse_pole_inline, parse_txt_upload, normalize_json_filename

print("TEST: parse_fat_inline(['A3','B2'])")
print(parse_fat_inline(['A3','B2']))

print("\nTEST: parse_pole_inline(['pole 73=3','ext 74=2'])")
print(parse_pole_inline(['pole 73=3','ext 74=2']))

sample_otdr = """FAT A01 | 0.194 | 1.000 | 0.196 | 1.340
FAT A02 | 0.201 | 0.980 | 0.203 | 1.310
"""

print("\nTEST: parse_txt_upload (otdr_cluster)")
print(parse_txt_upload(sample_otdr, 'otdr_cluster'))

print("\nTEST: normalize_json_filename('DUSUN BOGO RW 08 FDT-2')")
print(normalize_json_filename('DUSUN BOGO RW 08 FDT-2'))
