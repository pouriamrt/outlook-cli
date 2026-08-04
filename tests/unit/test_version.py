from typer.testing import CliRunner

from outlook_cli import __version__
from outlook_cli.cli import app


def test_version_command_prints_version() -> None:
    runner = CliRunner()
    result = runner.invoke(app, ["version"])
    assert result.exit_code == 0
    assert __version__ in result.stdout
