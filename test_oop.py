import unittest
from oop import Employee


class TestOop(unittest.TestCase):

    def test_class(self):
        emp1 = Employee('Test', 'User', 50000)
        self.assertIsInstance(emp1, Employee)
        self.assertEqual('Test', emp1.first)
        self.assertEqual('User', emp1.last)
        self.assertEqual(50000, emp1.pay)

    def test_fullname(self):
        emp1 = Employee('Test', 'User', 50000)
        self.assertEqual('Test User', emp1.fullname())


