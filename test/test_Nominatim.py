from src.NominatimIntegration import NominatimIntegration

import unittest


class TestNominatim(unittest.TestCase):

    def test_query_structured(self):
        result = NominatimIntegration.query_structured(amenity='UALR', street='2801 South University Avenue',
                                                       city='Little Rock', county='Pulaski', state='AR',
                                                       postalcode='72204')
        self.assertIsNotNone(result)
        self.assertIsInstance(result, dict)
        print(result)


if __name__ == '__main__':
    unittest.main()
