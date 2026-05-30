import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
MAINTENANCE = (ROOT / "PCMasterclass-Maintenance.ps1").read_text()


class WindowsIDriveOutlookBackupAuditTests(unittest.TestCase):
    def test_outlook_pst_discovery_checks_documents_onedrive_documents_and_appdata(self):
        self.assertIn('OutlookPstFiles', MAINTENANCE)
        self.assertIn('Documents\\*.pst', MAINTENANCE)
        self.assertIn('OneDrive\\Documents\\*.pst', MAINTENANCE)
        self.assertIn('AppData\\Local\\Microsoft\\Outlook\\*.pst', MAINTENANCE)

    def test_idrive_backup_set_audit_reports_pst_coverage_separately_from_log_success(self):
        self.assertIn('BackupSetCoverage', MAINTENANCE)
        self.assertIn('PST files found and backup set appears to include their folder', MAINTENANCE)
        self.assertIn('PST files found but no matching iDrive backup-set path found', MAINTENANCE)
        self.assertIn('Unable to verify iDrive backup-set coverage for Outlook PST files', MAINTENANCE)

    def test_onedrive_documents_pst_gets_specific_warning(self):
        self.assertIn('OneDrive-hosted PST files are unreliable', MAINTENANCE)
        self.assertIn('Move PST out of OneDrive and ensure iDrive backs it up', MAINTENANCE)

    def test_outlook_backup_audit_is_in_html_and_email_output(self):
        self.assertIn('iDrive Backup Set Coverage', MAINTENANCE)
        self.assertIn('iDRIVE BACKUP SET COVERAGE', MAINTENANCE)
        self.assertIn('BackupSetMatched', MAINTENANCE)

    def test_backup_set_coverage_notes_are_flattened_before_joining(self):
        self.assertIn('$allCoverageNotes = @($coverage)', MAINTENANCE)
        self.assertIn('$allCoverageNotes += $coverageNotes', MAINTENANCE)
        self.assertNotIn('($coverage, $coverageNotes)', MAINTENANCE)

    def test_standard_user_backup_folders_are_checked(self):
        self.assertIn('StandardUserFoldersCoverage', MAINTENANCE)
        for folder in ['Desktop', 'Documents', 'Downloads', 'Pictures', 'Music', 'Videos']:
            self.assertIn(f"'{folder}'", MAINTENANCE)
        self.assertIn('MissingStandardUserFolders', MAINTENANCE)
        self.assertIn('Standard user folders missing from iDrive backup set', MAINTENANCE)

    def test_outlook_roamcache_is_always_checked_for_outlook_users(self):
        self.assertIn('RoamCache', MAINTENANCE)
        self.assertIn('Stream_Autocomplete', MAINTENANCE)
        self.assertIn('Outlook RoamCache missing from iDrive backup set', MAINTENANCE)

    def test_root_drive_candidate_folders_are_reported_for_review(self):
        self.assertIn('RootDriveCandidateFolders', MAINTENANCE)
        self.assertIn('Previous PC backup', MAINTENANCE)
        self.assertIn('Root-level candidate folders found; consider adding to iDrive backup set', MAINTENANCE)
        self.assertIn('when in doubt, include due to iDrive generous storage quotas', MAINTENANCE)


if __name__ == "__main__":
    unittest.main()
