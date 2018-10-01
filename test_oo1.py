import unittest
import oo1


class TestOop (unittest.TestCase):

    def test_person(self):
        lar = oo1.Person('Larry', 'Shumlich', 29)
        lor = oo1.Person('Lorraine', 'Tkachyk', 27)
        lor.city = 'Calgary'

        self.assertEqual(27, lor.age)
        self.assertEqual(29, lar.age)
        self.assertEqual('Lorraine Tkachyk', lor.name)
        self.assertEqual('Larry Shumlich', lar.name)

        lar.first_name = 'Lars'
        self.assertEqual('Lars', lar.first_name)
        self.assertEqual('Lars Shumlich', lar.name)

        lar.last_name = 'Shumy'
        self.assertEqual('Shumy', lar.last_name)
        self.assertEqual('Lars Shumy', lar.name)

        lar.birthday()
        self.assertEqual(30, lar.age)

        lar.birthday()
        lar.birthday()
        lar.birthday()
        self.assertEqual(33, lar.age)
        self.assertEqual(27, lor.age)

