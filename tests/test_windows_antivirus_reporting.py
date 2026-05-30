import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
MAINTENANCE = (ROOT / "PCMasterclass-Maintenance.ps1").read_text()


class WindowsAntivirusReportingTests(unittest.TestCase):
    def test_report_has_dedicated_antivirus_inventory_section(self):
        self.assertIn('Antivirus Inventory / Endpoint Protection', MAINTENANCE)
        self.assertIn('SecurityCenter2Products', MAINTENANCE)
        self.assertIn('ProtectionSummary', MAINTENANCE)
        self.assertIn('RecommendedAction', MAINTENANCE)

    def test_malwarebytes_product_type_distinguishes_endpoint_premium_and_free(self):
        self.assertIn('Endpoint Protection', MAINTENANCE)
        self.assertIn('Premium - upgrade opportunity', MAINTENANCE)
        self.assertIn('Free - upgrade recommended', MAINTENANCE)
        self.assertIn('Consumer/Other - review before renewal', MAINTENANCE)

    def test_antivirus_summary_classifies_defender_third_party_and_missing_av(self):
        self.assertIn('Protected by PCMC-managed Malwarebytes Endpoint Protection', MAINTENANCE)
        self.assertIn('Protected by third-party antivirus', MAINTENANCE)
        self.assertIn('Protected by Microsoft Defender only', MAINTENANCE)
        self.assertIn('WARNING - No active antivirus detected', MAINTENANCE)
        self.assertIn('Multiple active antivirus products detected', MAINTENANCE)


if __name__ == "__main__":
    unittest.main()
