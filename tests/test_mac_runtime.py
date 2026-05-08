import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
INSTALL = (ROOT / "mac-maintenance" / "install.sh").read_text()
PLIST = (ROOT / "mac-maintenance" / "com.pcmasterclass.maintenance.agent.plist").read_text()
SCRIPT = (ROOT / "mac-maintenance" / "pcm_mac_maintenance.py").read_text()


class MacRuntimeInstallerTests(unittest.TestCase):
    def test_installer_defines_bundled_python_runtime(self):
        self.assertIn("python-build-standalone", INSTALL)
        self.assertIn("PYTHON_RUNTIME_DIR", INSTALL)
        self.assertIn("PYTHON_BIN", INSTALL)
        self.assertIn("aarch64-apple-darwin-install_only_stripped.tar.gz", INSTALL)

    def test_installer_does_not_invoke_apple_python_stub_for_runtime_work(self):
        forbidden = [
            'python3 "$INSTALL_DIR/$SCRIPT_NAME"',
            '/usr/bin/python3 "$INSTALL_DIR/$SCRIPT_NAME"',
            '<string>/usr/bin/python3</string>',
        ]
        for needle in forbidden:
            with self.subTest(needle=needle):
                self.assertNotIn(needle, INSTALL)

    def test_launchagent_uses_bundled_python_bin(self):
        self.assertIn('<string>$PYTHON_BIN</string>', INSTALL)
        self.assertIn('<string>$INSTALL_DIR/$SCRIPT_NAME</string>', INSTALL)

    def test_installer_uses_bundled_python_for_credentials_and_initial_scan(self):
        self.assertIn('"$PYTHON_BIN" "$INSTALL_DIR/$SCRIPT_NAME" --save-credential', INSTALL)
        self.assertIn('"$PYTHON_BIN" "$INSTALL_DIR/$SCRIPT_NAME" \\', INSTALL)
    def test_static_launchagent_template_uses_bundled_python(self):
        self.assertNotIn('<string>/usr/bin/python3</string>', PLIST)
        self.assertIn('/Library/PCMasterClass/python-runtime/python/bin/python3', PLIST)
    def test_smtp_app_password_spaces_are_normalized(self):
        self.assertIn('password = "".join(str(password).split())', SCRIPT)
        self.assertIn('args.smtp_password = "".join(args.smtp_password.split())', SCRIPT)
        self.assertIn('SMTP_PASS="${SMTP_PASS//[[:space:]]/}"', INSTALL)
    def test_browser_extension_inventory_handles_safari_permission_errors(self):
        self.assertIn('except PermissionError:', SCRIPT)
        self.assertIn('Browser extension inventory should never prevent report generation/email', SCRIPT)
        self.assertIn('Safari extensions inaccessible due to macOS privacy permissions', SCRIPT)


if __name__ == "__main__":
    unittest.main()
