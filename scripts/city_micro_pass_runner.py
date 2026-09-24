from importlib.util import module_from_spec, spec_from_file_location
from pathlib import Path

source = Path(__file__).with_name("city_micro_pass.py")
spec = spec_from_file_location("city_micro_pass", source)
if spec is None or spec.loader is None:
    raise RuntimeError(f"Cannot load {source}")
module = module_from_spec(spec)
spec.loader.exec_module(module)

# The canonical city key format is country::city. Keep the high-value city
# revalidation list on the same key format used by city_key().
module.PASS_FORMAT_VERSION = 13
module.FORCED_CITY_MIN_VERSION = {
    "poland::bydgoszcz": 13,
    "poland::warsaw": 13,
    "poland::łódź": 13,
    "germany::berlin": 13,
}

module.main()
