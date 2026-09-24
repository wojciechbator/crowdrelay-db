import scripts.city_micro_pass as m

# The canonical city key format is country::city. Keep the high-value city
# revalidation list on the same key format used by city_key().
m.PASS_FORMAT_VERSION = 6
m.FORCED_CITY_MIN_VERSION = {
    "poland::bydgoszcz": 6,
    "poland::warsaw": 6,
    "poland::łódź": 6,
    "germany::berlin": 6,
}

m.main()
