import json

from typer.testing import CliRunner

from outlook_cli.cli import app

runner = CliRunner()


def test_json_schema_mail_list() -> None:
    result = runner.invoke(app, ["--json-schema", "mail.list"])
    assert result.exit_code == 0
    schema = json.loads(result.stdout)
    assert schema["$schema"].startswith("https://json-schema.org/")
    assert schema["properties"]["items"]["type"] == "array"


def test_json_schema_unknown_emits_64() -> None:
    result = runner.invoke(app, ["--json-schema", "nope.bogus"])
    assert result.exit_code == 64


def test_version_command_prints_version_and_graph() -> None:
    result = runner.invoke(app, ["version"])
    assert result.exit_code == 0
    assert "outlook-cli" in result.stdout
    assert "graph" in result.stdout.lower()
