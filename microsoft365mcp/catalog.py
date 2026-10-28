"""Read and write tools on the Claude Microsoft 365 connector.

Names follow the connector security guide. Anything outside the read set is
hidden, so a write tool added upstream later stays disabled until it is
reviewed and added here.
"""

CONNECTOR_URL = "https://microsoft365.mcp.claude.com/mcp"

READ_TOOLS: frozenset[str] = frozenset(
    {
        "chat_message_search",
        "find_meeting_availability",
        "get_me",
        "outlook_calendar_search",
        "outlook_email_search",
        "outlook_find_available_time",
        "read_resource",
        "search_people",
        "sharepoint_folder_search",
        "sharepoint_search",
        "teams_list_channel_messages",
        "teams_list_channels",
        "teams_list_chats",
        "teams_list_teams",
    }
)

WRITE_TOOLS: frozenset[str] = frozenset(
    {
        "outlook_batch_delete_messages",
        "outlook_batch_modify_labels",
        "outlook_create_draft",
        "outlook_create_event",
        "outlook_create_filter",
        "outlook_create_label",
        "outlook_create_reply_all_draft",
        "outlook_create_reply_draft",
        "outlook_delete_draft",
        "outlook_delete_event",
        "outlook_delete_filter",
        "outlook_delete_label",
        "outlook_forward_mail",
        "outlook_modify_labels",
        "outlook_modify_thread_labels",
        "outlook_respond_to_event",
        "outlook_send_draft",
        "outlook_send_mail",
        "outlook_set_vacation",
        "outlook_trash_thread",
        "outlook_untrash_thread",
        "outlook_update_draft",
        "outlook_update_event",
        "outlook_update_label",
        "sharepoint_copy_item",
        "sharepoint_create_folder",
        "sharepoint_delete_item",
        "sharepoint_move_item",
        "sharepoint_rename_item",
        "sharepoint_update_file",
        "sharepoint_upload_file",
        "teams_create_chat",
        "teams_reply_channel_message",
        "teams_send_channel_message",
        "teams_send_chat_message",
    }
)

if not READ_TOOLS.isdisjoint(WRITE_TOOLS):
    raise RuntimeError("A Microsoft 365 tool is listed as both read and write.")
