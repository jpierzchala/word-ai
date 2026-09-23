import importlib.util
import json
from pathlib import Path
from types import SimpleNamespace
import unittest
from unittest.mock import patch

spec = importlib.util.spec_from_file_location("secure_launcher", Path(__file__).resolve().parents[1] / "scripts/start_secure_mcp.py")
launcher = importlib.util.module_from_spec(spec)
spec.loader.exec_module(launcher)


class LauncherGuards(unittest.TestCase):
    def test_missing_container_is_not_created(self):
        with patch.object(launcher, "engine_ready", return_value=True), patch.object(launcher, "run", return_value=SimpleNamespace(returncode=1, stdout="")) as run:
            with self.assertRaisesRegex(RuntimeError, "container not found"):
                launcher.ensure_ready("docker")
            self.assertEqual(run.call_args_list[0].args[0], ["docker", "inspect", "--type", "container", "word-ai-secure"])
            self.assertEqual(run.call_count, 1)

    def test_unrelated_container_is_never_started_or_attached(self):
        unexpected = [{"Config": {"Labels": {"com.docker.compose.project": "another-project"}}, "State": {"Running": False}}]
        with patch.object(launcher, "engine_ready", return_value=True), patch.object(launcher, "run", return_value=SimpleNamespace(returncode=0, stdout=json.dumps(unexpected))) as run:
            with self.assertRaisesRegex(RuntimeError, "unexpected deployment"):
                launcher.ensure_ready("docker")
            self.assertEqual(run.call_count, 1)

    def test_failed_engine_start_has_a_bounded_timeout(self):
        with patch.object(launcher, "engine_ready", return_value=False), patch.object(launcher, "start_rancher") as start, patch.object(launcher.time, "monotonic", side_effect=[0, 181]), patch.object(launcher.time, "sleep") as sleep:
            with self.assertRaisesRegex(RuntimeError, "180 seconds"):
                launcher.ensure_ready("docker")
            start.assert_called_once()
            sleep.assert_not_called()


if __name__ == "__main__":
    unittest.main()
