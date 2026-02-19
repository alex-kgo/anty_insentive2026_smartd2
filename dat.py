from pathlib import Path

p = Path(r"C:\Users\alex\AppData\Local\Temp\{72167FFA-BC8F-4843-9C69-C01688437EEE} - OProcSessId.dat")  # 파일 경로로 바꿔줘
b = p.read_bytes()[:16]
print("HEAD:", b)
print("HEX :", b.hex())

# 간단한 타입 추정
if b.startswith(b"PK\x03\x04"):
    print("-> ZIP 계열(=xlsx 같은 zip 컨테이너 가능성)")
elif b.startswith(b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1"):
    print("-> OLE Compound(=xls 같은 구형 포맷/오피스 컨테이너 가능성)")
else:
    print("-> 일반 바이너리/세션파일 가능성이 큼")
