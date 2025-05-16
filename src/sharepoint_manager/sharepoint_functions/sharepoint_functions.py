import io
import os
from typing import List, Union

from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.attachments.attachmentfile_collection import (
    AttachmentFileCollection,
)
from office365.sharepoint.attachments.attachmentfile_creation_information import (
    AttachmentfileCreationInformation,
)
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file_collection import FileCollection
from office365.sharepoint.folders.folder_collection import FolderCollection
from office365.sharepoint.listitems.listitem import ListItem
from office365.sharepoint.lists.list import List as ListSharepoint
from office365.sharepoint.lists.list_creation_information import ListCreationInformation
from office365.sharepoint.lists.list_template_type import ListTemplateType
from pydantic import BaseModel


class Information(BaseModel):
    """
    Defines a class called "Information" which is a subclass of
    "BaseModel". The "Information" class is used to load information
    into a specific column.

    Args:
        column (str): Name of the column to which the information will
        be loaded.

        value (str): Data to be loaded to the column.
    """

    column: str
    value: str


class SharepointErrors(Exception):
    """
    Defines a custom exception class called "SharepointErrors" that
    inherits from the built-in "Exception" class. The purpose of this
    class is to handle and represent errors related to SharePoint
    operations or functionalities.
    """

    def __init__(self, message):
        self.message = message
        super().__init__(self.message)


def get_connection_sharepoint(url: str, username: str, password: str) -> ClientContext:
    """
    This function establish a connection to a SharePoint site using the
    provided URL, username, and password.

    Args:
        url (str): sharepoint url

        username (str): username like something@anla.gov.co

        password (str): password of the username

    Returns:
        ClientContext: object representing a client context in a
        SharePoint environment. A ClientContext typically provides
        access to SharePoint resources and operations.
    """

    ctx_auth = AuthenticationContext(url)
    if ctx_auth.acquire_token_for_user(username, password):
        ctx = ClientContext(url, ctx_auth)
        web = ctx.web
        ctx.load(web)
        ctx.execute_query()
        return ctx
    else:
        print(f"The connection to {url} failed")


def get_connection_list(ctx: ClientContext, list_name: str) -> ListSharepoint:
    """
    This function takes a SharePoint client context (ctx) and the name
    of a SharePoint list (list_name) as inputs. It then uses the client
    context to access the SharePoint web and retrieves the SharePoint
    list specified by its title. The function ultimately returns the
    SharePoint list object.

    Args:
        ctx (ClientContext): object representing a client context in a
        SharePoint environment. A ClientContext typically provides
        access to SharePoint resources and operations.

        list_name (str): it expects a string value representing the name
        of a SharePoint list.

    Returns:
        ListSharepoint: The function will return an object representing
        a SharePoint list.
    """

    return ctx.web.lists.get_by_title(list_name)


def get_connection_folder(ctx: ClientContext, folder_name: str) -> ListSharepoint:
    """
    In summary, the get_connection_folder function takes a SharePoint
    client context (ctx) and the server-relative URL of a SharePoint
    folder (folder_name) as inputs. It uses the client context to access
    the SharePoint web and retrieves the SharePoint folder specified by
    its server-relative URL. The function ultimately returns the
    SharePoint folder object.

     Args:
         ctx (ClientContext): object representing a client context in a
         SharePoint environment. A ClientContext typically provides
         access to SharePoint resources and operations.

         folder_name (str): it expects a string value representing the
         server-relative URL of a SharePoint folder.

     Returns:
         ListSharepoint: The function will return an object representing
         a SharePoint list.
    """

    return ctx.web.get_folder_by_server_relative_url(folder_name)


def get_info_folders(
    ctx: ClientContext, folder_sharepoint: ListSharepoint
) -> FolderCollection:
    """
    It retrieves the collection of folders within the specified
    SharePoint list by accessing the folders property. The folders
    collection is loaded into the client context, and the query is
    executed to fetch the data from the server.

    Args:
        ctx (ClientContext): object representing a client context in a
         SharePoint environment. A ClientContext typically provides
         access to SharePoint resources and operations.

        folder_sharepoint (ListSharepoint): It expects an object
        representing a SharePoint list.

    Returns:
        folders: The function then returns the retrieved folders
        collection.
    """

    folders = folder_sharepoint.folders
    ctx.load(folders)
    ctx.execute_query()

    return folders


def get_info_files(
    ctx: ClientContext, file_sharepoint: ListSharepoint
) -> FileCollection:
    """
    It retrieves the collection of files within the specified SharePoint
    list by accessing the files property. The files collection is loaded
    into the client context, and the query is executed to fetch the data
    from the server. The function then returns the retrieved files
    collection.

    Args:
        ctx (ClientContext): object representing a client context in a
         SharePoint environment. A ClientContext typically provides
         access to SharePoint resources and operations.

        file_sharepoint (ListSharepoint): It expects an object
        representing a SharePoint list.

    Returns:
        files: the function will return a collection of SharePoint
        files.
    """

    files = file_sharepoint.files
    ctx.load(files)
    ctx.execute_query()

    return files


def get_info_list(
    ctx: ClientContext,
    list_sharepoint: ListSharepoint,
    filters: List[Information] = None,
) -> ListItem:
    """
    The get_info_list function takes a SharePoint client context (ctx),
    a SharePoint list object (list_sharepoint), and an optional list of
    filters (filters). It retrieves the items from the specified
    SharePoint list, applies filters if provided, and returns the
    resulting items.

    Args:
        ctx (ClientContext): object representing a client context in a
         SharePoint environment. A ClientContext typically provides
         access to SharePoint resources and operations.

        list_sharepoint (ListSharepoint): It expects an object
        representing a SharePoint list.

        filters (List[Information], optional):  It represents
        information used for filtering the items in the list.

    Returns:
        ListItem: The function returns either the filtered items list or
        the original items collection, depending on the presence of
        filters.
    """

    items = list_sharepoint.items
    ctx.load(items)
    ctx.execute_query()
    items = (
        [
            item
            for item in items
            if all(
                str(item.properties[query.column]) == query.value for query in filters
            )
        ]
        if filters is not None
        else items
    )

    return items


def get_email_user(ctx: ClientContext, professional_id: int) -> str:
    """
    The get_email_user function takes a SharePoint client context (ctx)
    and a professional ID (professional_id). It retrieves the SharePoint
    site users, searches for a user with a matching ID, and returns the
    email address of that user if found. If no match is found,
    it returns None.

    Args:
        ctx (ClientContext): object representing a client context in a
         SharePoint environment. A ClientContext typically provides
         access to SharePoint resources and operations.

        professional_id (int): Represents the ID of the professional
        within SharePoint.

    Returns:
        str: Sharepoint user mail
    """
    users = ctx.web.site_users
    ctx.load(users)
    ctx.execute_query()
    for user in users:
        if user.properties["Id"] == professional_id:
            return user.properties["UserPrincipalName"]
    return None


def extract_attachment_files(
    ctx: ClientContext, item: ListItem
) -> AttachmentFileCollection:
    """
    The extract_attachment_files function takes a SharePoint client
    context (ctx) and a SharePoint item (item). It retrieves the
    attachment files associated with the item by loading and executing
    the necessary query, and returns the collection of attachment files.

    Args:
        ctx (ClientContext): object representing a client context in a
         SharePoint environment. A ClientContext typically provides
         access to SharePoint resources and operations.

        item (ListItem): Represents a SharePoint item from which the
        attachment files are to be extracted.

    Returns:
        AttachmentFileCollection: Represents the attachment files
        associated with the SharePoint item.
    """
    attachment_files = item.attachment_files
    ctx.load(attachment_files)
    ctx.execute_query()
    return attachment_files


def update_list(
    ctx: ClientContext,
    list_sharepoint: ListSharepoint,
    id_item: int,
    update_info: List[Information],
) -> None:
    """
    The update_list function takes a SharePoint client context (ctx), a
    SharePoint list object (list_sharepoint), an item ID (id_item), and
    a list of information to update (update_info). It retrieves the
    specified item from the SharePoint list, updates the specified
    information on the item, and saves the changes back to the server.
    If the item is not found, an exception is raised.

    Args:
        ctx (ClientContext): object representing a client context in a
         SharePoint environment. A ClientContext typically provides
         access to SharePoint resources and operations.

        list_sharepoint (ListSharepoint): It expects an object
        representing a SharePoint list.

        id_item (int): Represents the ID of the item in the SharePoint
        list that needs to be updated.

        update_info (List[Information]): Represents the list of
        information to update on the SharePoint item.
    """

    items = list_sharepoint.items
    ctx.load(items)
    ctx.execute_query()
    item_selection = [item for item in items if item.properties["Id"] == id_item]
    if item_selection:
        item = item_selection[0]
        for information in update_info:
            item.set_property(information.column, information.value)
            item.update()
            ctx.execute_query()
    else:
        raise SharepointErrors(
            f"The item {id_item} do nor exist in the sharepoint list"
        )


def create_list(ctx: ClientContext, list_name: str) -> None:
    """
    The create_list function takes a SharePoint client context (ctx) and
    a list name (list_name). It creates a new SharePoint list with the
    specified name, using the generic list template and allowing content
    types.

    Args:
        ctx (ClientContext): object representing a client context in a
         SharePoint environment. A ClientContext typically provides
         access to SharePoint resources and operations.

        list_name (str): Represents the name of the new SharePoint list.
    """
    list_creation = ListCreationInformation()
    list_creation.BaseTemplate = ListTemplateType.GenericList
    list_creation.Title = list_name
    list_creation.AllowContentTypes = True
    ctx.web.lists.add(list_creation).execute_query()


def add_new_item(
    ctx: ClientContext,
    list_sharepoint: ListSharepoint,
    information: dict,
    attachment: str = None,
) -> None:
    """
    The add_new_item function takes a SharePoint client context (ctx), a
    SharePoint list object (list_sharepoint), a dictionary of
    information for the new item (information), and an optional
    attachment file path (attachment). It adds a new item to the
    SharePoint list using the provided information and saves it. If an
    attachment file is specified, it adds the attachment to the newly
    created item.

    Args:
        ctx (ClientContext): object representing a client context in a
         SharePoint environment. A ClientContext typically provides
         access to SharePoint resources and operations.

        list_sharepoint (ListSharepoint): It expects an object representing a
        SharePoint list.

        information (dict): Represents the information to be included
        in the new item. It is expected to be a dictionary where the
        keys represent the column names and the values represent the
        column values.

        attachment (str, optional): Path to an optional attachment file
        to be associated with the new item. It has a default value of
        None.
    """
    item = list_sharepoint.add_item(information)
    ctx.execute_query()
    if attachment:
        with open(attachment, "rb") as file:
            file_content = file.read()
        attachment_file_info = AttachmentfileCreationInformation(
            os.path.basename(attachment), file_content
        )
        item.attachment_files.add(attachment_file_info).execute_query()


def create_folder(ctx: ClientContext, relative_path: str, folder_name: str) -> None:
    """
    The create_folder function takes a SharePoint client context (ctx),
    a relative path (relative_path), and a folder name (folder_name).
    It creates a new folder in SharePoint by adding it to the folders
    collection of the SharePoint web using the specified relative path
    and folder name.

    Args:
        ctx (ClientContext): object representing a client context in a
         SharePoint environment. A ClientContext typically provides
         access to SharePoint resources and operations.

        relative_path (str): Relative path where the new folder will be
        created.

        folder_name (str): Name of the new folder to be created.
    """
    ctx.web.folders.add(f"{relative_path}/{folder_name}")
    ctx.execute_query()


def is_sharepoint_folder(
    ctx: ClientContext, relative_path: str, folder_name: str
) -> bool:
    """
    The is_sharepoint_folder function takes a SharePoint client
    context (ctx), a relative path of the parent folder (relative_path),
    and a target folder name (folder_name). It retrieves the sub-folders
    within the parent folder, checks if the target folder name exists
    within the sub-folders, and returns a boolean value indicating
    whether the target folder exists within the SharePoint site.

    Args:
        ctx (ClientContext): object representing a client context in a
         SharePoint environment. A ClientContext typically provides
         access to SharePoint resources and operations.

        relative_path (str): Relative path where the new folder will be
        created.

        folder_name (str): Name of the new folder to be created.

    Returns:
        bool: The function checks whether folder_name exists within the
        list_sub_folder list and returns True if it does, or False
        otherwise.
    """
    target_folder = ctx.web.get_folder_by_server_relative_url(relative_path)
    sub_folders = target_folder.folders
    ctx.load(sub_folders)
    ctx.execute_query()
    list_sub_folder = [
        sub_folders[index].properties["Name"] for index in range(0, len(sub_folders))
    ]
    return folder_name in list_sub_folder


def upload_file_sharepoint(
    ctx: ClientContext,
    relative_path: str,
    file_upload: Union[str, io.BytesIO],
    file_name: str = None,
) -> None:
    """
    The upload_file_sharepoint function takes a SharePoint client
    context (ctx), a relative path of the target folder (relative_path),
    the file to be uploaded (file_upload), and an optional file name
    (file_name). It uploads the file to the specified folder in
    SharePoint, either using the provided file name or extracting it
    from the file path.

    Args:
        ctx (ClientContext): object representing a client context in a
         SharePoint environment. A ClientContext typically provides
         access to SharePoint resources and operations.

        relative_path (str): Relative path of the target folder where
        the file will be uploaded.

        file_upload (Union[str, io.BytesIO]): The file to be uploaded.
        It can be either a file path (str) or a BytesIO object.

        file_name (str, optional): Desired name of the file in
        SharePoint. It has a default value of None.
    """
    target_folder = ctx.web.get_folder_by_server_relative_url(relative_path)
    if isinstance(file_upload, str):
        with open(file_upload, "rb") as file:
            file_content = file.read()
    elif isinstance(file_upload, io.BytesIO):
        file_content = file_upload.getvalue()
    else:
        raise TypeError("El argumento file_upload debe ser una ruta de archivo o un objeto io.BytesIO")
    if file_name is None:
        file_name = file_upload.split("\\")[-1]
    target_folder.upload_file(file_name, file_content).execute_query()


def delete_item_list(ctx: ClientContext, item: ListItem) -> None:
    """
    Deletes a specific SharePoint item.

    Args:
        ctx (ClientContext): SharePoint client context.
        item (ListItem): Item to be deleted.
    """
    item.delete_object()
    ctx.execute_query()


def delete_folder_with_contents(ctx: ClientContext, folder_path: str) -> None:
    """
    Recursively deletes all files and subfolders in a SharePoint folder,
    and then deletes the folder itself.

    Args:
        ctx (ClientContext): SharePoint client context.
        folder_path (str): Server-relative path of the folder to delete.
    """
    folder = ctx.web.get_folder_by_server_relative_url(folder_path)
    ctx.load(folder)
    ctx.execute_query()

    # Delete files
    files = folder.files
    ctx.load(files)
    ctx.execute_query()
    for file in files:
        file.delete_object()
    ctx.execute_query()

    # Delete subfolders recursively
    subfolders = folder.folders
    ctx.load(subfolders)
    ctx.execute_query()
    for subfolder in subfolders:
        delete_folder_with_contents(ctx, subfolder.properties["ServerRelativeUrl"])

    # Delete the folder itself
    folder.delete_object()
    ctx.execute_query()

