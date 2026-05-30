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

    def test_unwanted_security_scanware_is_detected_and_reported(self):
        self.assertIn('UnwantedSecuritySoftware', MAINTENANCE)
        self.assertIn('McAfee Security Scan Plus', MAINTENANCE)
        self.assertIn('Norton Security Scan', MAINTENANCE)
        self.assertIn('Security scanware / unwanted AV-adjacent software', MAINTENANCE)
        self.assertIn('Remove useless security scanware such as McAfee Security Scan Plus', MAINTENANCE)

    def test_unwanted_security_scanware_is_in_email_and_html_sections(self):
        self.assertIn('UNWANTED SECURITY SOFTWARE', MAINTENANCE)
        self.assertIn('Unwanted security software detected', MAINTENANCE)
        self.assertIn('if ($Results.AntivirusInventory.UnwantedSecuritySoftware', MAINTENANCE)

    def test_email_subject_preserves_existing_surname_firstname_comma(self):
        self.assertIn('preserve the comma rather than adding a second one', MAINTENANCE)
        self.assertIn("$ClientName -match '^\\s*([^,]+),\\s*(.+?)\\s*$'", MAINTENANCE)
        self.assertIn('$($Matches[1].Trim().ToUpper()), $($Matches[2].Trim())', MAINTENANCE)


if __name__ == "__main__":
    unittest.main()
