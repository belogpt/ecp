import os
import sys
import unittest

from signer_cadescom import SignerCadescomError, _ensure_com_available, list_certificates


class SignerCadescomTests(unittest.TestCase):
    def test_com_available(self):
        if sys.platform != "win32":
            self.skipTest("Тест доступен только в Windows окружении")
        try:
            _ensure_com_available()
        except SignerCadescomError as exc:  # pragma: no cover - зависит от окружения
            self.fail(f"CAdESCOM недоступен: {exc}")

    def test_certificates_readable(self):
        if sys.platform != "win32":
            self.skipTest("Тест доступен только в Windows окружении")
        certs = list_certificates()
        self.assertIsInstance(certs, list)
        # Не проверяем содержимое: зависит от установленного CSP и наличия сертификатов.

    @unittest.skip("Manual: требуется подключенный токен/сертификат и рабочий CSP")
    def test_sign_file_manual(self):
        sample = os.path.abspath(__file__)
        output = sample + ".p7s"
        # Этот тест запускается вручную в Windows для проверки создания подписи через CAdESCOM.
        from signer_cadescom import sign_file

        sign_file(sample, output_path=output)
        self.assertTrue(os.path.exists(output))
        os.remove(output)


if __name__ == "__main__":
    unittest.main()
