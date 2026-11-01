import unittest

from routebot.addresses import address_variants, clean_addr, normalize_street, split_contact


class AddressesTest(unittest.TestCase):
    def test_clean_addr_strips_city_and_fixes_spacing(self):
        self.assertEqual(clean_addr("г.Астана, ул.Толе би 46"), "ул. Толе би 46")
        self.assertEqual(clean_addr("Нур-Султан,  пр.Абая 48"), "пр. Абая 48")
        self.assertEqual(clean_addr("⁠ул. Титова 14/2"), "ул. Титова 14/2")

    def test_split_contact(self):
        self.assertEqual(
            split_contact("Жамал Манас +7 702 222 0723", ""), ("Жамал Манас", "+7 702 222 0723")
        )
        self.assertEqual(split_contact("Мира", "+7 777"), ("Мира", "+7 777"))
        self.assertEqual(split_contact("Гульжанат", ""), ("Гульжанат", ""))

    def test_letter_and_numeric_streets(self):
        self.assertEqual(normalize_street("ул. Е 652, 22"), "улица E-652, 22")
        self.assertIn("улица 23-16", normalize_street("ул. 23-16 дом 39"))
        self.assertEqual(normalize_street("ул. Аль-Фараби 35, 14"), "проспект Аль-Фараби 35")

    def test_variants_are_unique_and_include_base_house(self):
        v = address_variants("ул. Байтурсынова 39/2, ЖК Сатурн")
        self.assertEqual(len(v), len(set(v)))
        self.assertIn("улица Ахмета Байтурсынова 39", v)
        self.assertIn("ЖК Сатурн", v)

    def test_house_number_of_numbered_road_is_kept(self):
        # Раньше «, 22» отрезалось как номер квартиры — пин вставал на середину улицы
        self.assertEqual(normalize_street("ул. Е 652, 22"), "улица E-652, 22")
        self.assertEqual(normalize_street("ул. 23-16 дом 39"), "улица 23-16, 39")

    def test_landmarks_and_district_are_split_off(self):
        self.assertIn("шоссе Алаш 31/1", address_variants("Шоссе Алаш 31/1 Рынок Шарын"))
        v = address_variants(clean_addr("Город Нур-Султан, Алматинский район, ул. 23-16 дом 39"))
        self.assertIn("улица 23-16, 39", v)
        self.assertIn("Алматинский район, улица 23-16, 39", v)
        self.assertFalse(any("Нур-Султан" in x for x in v))
        # «23-16, 39» без «улица» OSM путает с «39-й улицей» в другом районе
        self.assertNotIn("23-16, 39", v)
        self.assertNotIn("E-652, 22", address_variants("ул. Е 652, 22"))

    def test_name_search_only_when_allowed(self):
        self.assertNotIn("Султан", address_variants("ул. Абая 1", "Султан GS"))
        self.assertIn("Султан", address_variants("ул. Абая 1", "Султан GS", allow_name_search=True))


if __name__ == "__main__":
    unittest.main()
