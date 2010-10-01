# file openpyxl/tests/test_password_hash.py

from legacy_openpyxl.tests.helper import BaseTestCase

from legacy_openpyxl.shared.password_hasher import hash_password

from legacy_openpyxl.worksheet import SheetProtection

class TestPasswordHasher(BaseTestCase):

    def test_hasher(self):

         self.assertEqual('CBEB', hash_password('test'))


    def test_sheet_protection(self):

        p = SheetProtection()

        p.password = 'test'

        self.assertEqual('CBEB', p.password)
