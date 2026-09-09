import json,sys
from core import generate_single_image
spec=json.load(open(sys.argv[1],encoding='utf8'))
path=generate_single_image(spec['drawing'],spec['save_dir'])
print(json.dumps({'path':path},ensure_ascii=False),flush=True)
