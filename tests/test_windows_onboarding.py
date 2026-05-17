import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
ONBOARDING = (ROOT / "Onboarding-PCMasterclass.ps1").read_text()


class WindowsOnboardingTests(unittest.TestCase):
    def test_deployment_test_email_subject_matches_maintenance_report_filter(self):
        self.assertIn('"[TEST] Maintenance Report Delivery Test - $computerName"', ONBOARDING)
        self.assertNotIn('"PC Masterclass - Deployment Test - $computerName"', ONBOARDING)

    def test_scheduled_maintenance_defaults_to_initial_run_today(self):
        self.assertIn('$todayStr = (Get-Date).ToString("dd/MM/yyyy")', ONBOARDING)
        self.assertIn('Start date dd/MM/yyyy (press Enter for today, $todayStr)', ONBOARDING)
        self.assertIn('$startDate = (Get-Date).Date', ONBOARDING)
        self.assertIn('Starting first maintenance run in background', ONBOARDING)

    def test_documentation_says_onboarding_runs_initial_maintenance_by_default(self):
        self.assertIn('Runs the first maintenance scan immediately after schedule creation by default', ONBOARDING)
        self.assertNotIn('It does not run the maintenance script (you decide when to run it)', ONBOARDING)


if __name__ == "__main__":
    unittest.main()
