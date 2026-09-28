from importlib.util import module_from_spec, spec_from_file_location
from pathlib import Path

source = Path(__file__).with_name("city_micro_pass.py")
spec = spec_from_file_location("city_micro_pass", source)
if spec is None or spec.loader is None:
    raise RuntimeError(f"Cannot load {source}")
module = module_from_spec(spec)
spec.loader.exec_module(module)

module.main()
