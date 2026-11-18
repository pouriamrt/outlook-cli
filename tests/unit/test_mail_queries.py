from datetime import UTC, datetime
from unittest.mock import MagicMock

from hypothesis import given
from hypothesis import strategies as st

from outlook_cli.graph.mail import MailListFilters, build_list_params, list_messages


def test_build_params_default_inbox_unread() -> None:
    params = build_list_params(MailListFilters(unread=True, top=25))
    assert params["$top"] == 25
    assert "isRead eq false" in params["$filter"]
    assert params["$orderby"] == "receivedDateTime desc"
    assert "id,subject" in params["$select"]


def test_build_params_combines_unread_and_from() -> None:
    params = build_list_params(MailListFilters(unread=True, from_addr="alice@example.com", top=10))
    assert "isRead eq false" in params["$filter"]
    assert "from/emailAddress/address eq 'alice@example.com'" in params["$filter"]


def test_build_params_includes_since_filter() -> None:
    since = datetime(2026, 5, 20, 12, 0, 0, tzinfo=UTC)
    params = build_list_params(MailListFilters(since=since, top=25))
    assert "receivedDateTime ge 2026-05-20T12:00:00" in params["$filter"]


def test_build_params_from_and_since_puts_received_first() -> None:
    # Regression: Exchange returned 400 InefficientFilter for this exact combination.
    since = datetime(2026, 5, 20, 12, 0, 0, tzinfo=UTC)
    params = build_list_params(MailListFilters(from_addr="alice@example.com", since=since))
    assert params["$filter"] == (
        "receivedDateTime ge 2026-05-20T12:00:00Z"
        " and from/emailAddress/address eq 'alice@example.com'"
    )


def test_build_params_no_filters_has_no_filter_key() -> None:
    assert "$filter" not in build_list_params(MailListFilters())


@given(
    unread=st.booleans(),
    flagged=st.booleans(),
    from_addr=st.none() | st.text(min_size=1),
    subject=st.none() | st.text(min_size=1),
    since=st.none() | st.datetimes(timezones=st.just(UTC)),
)
def test_build_params_orderby_property_leads_filter(
    unread: bool,
    flagged: bool,
    from_addr: str | None,
    subject: str | None,
    since: datetime | None,
) -> None:
    filters = MailListFilters(
        unread=unread, flagged=flagged, from_addr=from_addr, subject=subject, since=since
    )
    params = build_list_params(filters)
    if "$filter" in params:
        assert params["$filter"].startswith("receivedDateTime ge ")
        assert params["$filter"].count("receivedDateTime") == 1


def test_build_params_skip_when_provided() -> None:
    params = build_list_params(MailListFilters(top=25, skip=50))
    assert params["$skip"] == 50


def test_list_messages_calls_graph_with_folder_path() -> None:
    client = MagicMock()
    client.get.return_value.json.return_value = {"value": []}
    list_messages(client, folder_id="inbox", filters=MailListFilters(top=5))
    args, kwargs = client.get.call_args
    assert "/me/mailFolders/inbox/messages" in args[0]


def test_list_messages_returns_parsed_messages() -> None:
    client = MagicMock()
    client.get.return_value.json.return_value = {
        "value": [
            {
                "id": "MSG1",
                "subject": "hi",
                "isRead": False,
                "receivedDateTime": "2026-05-22T10:00:00Z",
                "from": {"emailAddress": {"name": "A", "address": "a@b"}},
            }
        ],
        "@odata.nextLink": "https://outlook.office.com/api/v2.0/me/messages?$skip=25",
    }
    result = list_messages(client, folder_id="inbox", filters=MailListFilters(top=5))
    assert len(result.items) == 1
    assert result.items[0].subject == "hi"
    assert result.next_link is not None
