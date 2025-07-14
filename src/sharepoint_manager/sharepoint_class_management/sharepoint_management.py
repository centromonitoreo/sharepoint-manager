import io
import os
from typing import List, Union

from office365.runtime.auth.authentication_context import AuthenticationContext
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

from sharepoint_manager.sharepoint_functions.sharepoint_funcions import (
    add_new_item,
    create_folder,
    create_list,
    get_connection_folder,
    get_connection_list,
    get_connection_sharepoint,
    get_email_user,
    get_info_list,
    is_sharepoint_folder,
    update_list,
    upload_file_sharepoint,
    get_info_folders,
    get_info_files,
    delete_item_list,
    delete_folder_with_contents,
)


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


class SharepointManagement:
    """Sharepoint connection and control class"""

    def __init__(self, username: str, password: str, url: str) -> None:
        """Constructor of the SharepointConnection class

        Args:
            username (str): Email of the sharepoint user.

            password (str): Outlook mail authentication password.

            url (str): Url of the sharepoint to which you want to make the connection.
        """
        self.username = username
        self.password = password
        self.url = url
        self.ctx = None
        self.list_conexion = {}
        self.folder_conexion = {}
        self.get_conexion_sharepoint()

    def get_conexion_sharepoint(self) -> None:
        """
        The get_conexion_sharepoint method is responsible for obtaining
        a connection to SharePoint by calling the
        get_connection_sharepoint function with the necessary parameters
        (self.url, self.username, self.password). The connection object
        returned by the function is then assigned to the ctx attribute
        of the class instance.
        """
        self.ctx = get_connection_sharepoint(self.url, self.username, self.password)

    def get_conexion_list(self, list_name) -> None:
        """
        The get_conexion_list method is responsible for connecting to a
        SharePoint list specified by list_name. It does so by calling
        the get_connection_list function with the SharePoint client
        context (self.ctx) and the list name as arguments. The result
        of the function call is then stored in the list_conexion
        attribute of the class instance, with list_name as the key.

        Args:
            list_name: Is the name of the SharePoint list to connect to.
        """
        self.list_conexion[list_name] = get_connection_list(self.ctx, list_name)

    def get_conexion_folder(self, folder_name) -> None:
        """
        The get_conexion_folder method is responsible for connecting to
        a SharePoint folder specified by folder_name. It does so by
        calling the get_connection_folder function with the SharePoint
        client context (self.ctx) and the folder name as arguments.
        The result of the function call is then stored in the
        folder_conexion attribute of the class instance, with
        folder_name as the key.

        Args:
            folder_name: is the name or identifier of the SharePoint
            folder to connect to.
        """
        self.folder_conexion[folder_name] = get_connection_folder(self.ctx, folder_name)

    def get_email_user(self, profesional_id: int) -> str:
        """
        the get_email_user method is method is a wrapper that delegates
        the task of retrieving the email of a SharePoint user to another
        function or method named get_email_user. The SharePoint client
        context (self.ctx) and the professional ID are passed as
        arguments to this function. The method simply returns the
        result obtained from calling the get_email_user function.

        Args:
            professional_id (int): Represents the ID of the professional
        within SharePoint.

        Returns:
            str: sharepoint user mail.
        """
        return get_email_user(self.ctx, profesional_id)

    def get_info_list(
        self,
        list_sharepoint: str,
        filters: List[Information] = None,
    ) -> dict:
        """
        The get_info_list method is a wrapper that delegates the task of
        extracting information from a SharePoint list to another
        function named get_info_list. The SharePoint client
        context (self.ctx), the connection to the
        list (self.list_conexion[list_sharepoint]), and optional filters
        are passed as arguments to this function. The method simply
        returns the result obtained from calling the get_info_list
        function.

        Args:
            list_sharepoint (str): Name of the SharePoint list.

            filters (List[Information], optional):  It represents
            information used for filtering the items in the list.

        Returns:
            ListItem: The function returns either the filtered items
            list or the original items collection, depending on the
            presence of filters.
        """
        return get_info_list(
            self.ctx,
            self.list_conexion[list_sharepoint],
            filters,
        )

    def get_info_folders(self, folder_name: str):
        """
        The get_info_folders method is a wrapper that delegates the task
        of obtaining information about folders in SharePoint to another
        function or method named get_info_folders. The SharePoint client
        context (self.ctx) and the connection to the folder
        (self.folder_conexion[folder_name]) are passed as arguments to
        this function. The method simply returns the result obtained
        from calling the get_info_folders function.

        Args:
            folder_name (str): Name of the folder name.

        Returns:
            The function then returns the retrieved folders collection.
        """

        return get_info_folders(self.ctx, self.folder_conexion[folder_name])

    def get_info_files(self, folder_name: str):
        """
        The get_info_files method is a wrapper that delegates the task
        of obtaining information about files in a SharePoint folder to
        another function or method named get_info_files. The SharePoint
        client context (self.ctx) and the connection to the folder
        (self.folder_conexion[folder_name]) are passed as arguments to
        this function. The method simply returns the result obtained
        from calling the get_info_files function.

        Args:
            folder_name (str): Name of the folder name.

        Returns:
            The function will return a collection of SharePoint files.
        """
        return get_info_files(self.ctx, self.folder_conexion[folder_name])

    def update_list(
        self, list_name: str, id_item: int, update_info: List[Information]
    ) -> None:
        """
        The update_list method is a wrapper that delegates the task of
        updating a SharePoint list to another function named
        update_list. The SharePoint client context (self.ctx), the
        connection to the list (self.list_conexion[list_name]), the
        item ID (id_item), and the update information (update_info)
        are passed as arguments to this function.

        Args:
            list_name (str): Name of the sharepoint list.

            id_item (int): Represents the ID of the item in the SharePoint
            list that needs to be updated.

            update_info (List[Information]): Represents the list of
            information to update on the SharePoint item.
        """
        update_list(
            self.ctx,
            self.list_conexion[list_name],
            id_item,
            update_info,
        )

    def add_new_item(
        self, list_name: str, information: dict, attachment: str = None
    ) -> None:
        """

        The add_new_item method is a wrapper that delegates the task of
        adding a new item to a SharePoint list to another function named
        add_new_item. The SharePoint client context (self.ctx), the
        connection to the list (self.list_conexion[list_name]), the
        information of the new item (information), and the attachment
        file path (attachment) are passed as arguments to this function.

        Args:
            list_name (str): Name of the sharepoint list.

            information (dict): Represents the information to be included
            in the new item. It is expected to be a dictionary where the
            keys represent the column names and the values represent the
            column values.

            attachment (str, optional): Path to an optional attachment file
            to be associated with the new item. It has a default value of
            None.
        """
        add_new_item(self.ctx, self.list_conexion[list_name], information, attachment)

    def create_folder(self, relative_path: str, folder_name: str) -> None:
        """
        The create_folder method is a wrapper that delegates the task of
        creating a folder in SharePoint to another function named
        create_folder. The SharePoint client context (self.ctx), the
        relative path of the folder (relative_path), and the folder
        name (folder_name) are passed as arguments to this function.

        Args:
            relative_path (str): Relative path where the new folder will
            be created.

            folder_name (str): Name of the new folder to be created.
        """
        create_folder(self.ctx, relative_path, folder_name)

    def is_sharepoint_folder(self, relative_path: str, folder_name: str) -> bool:
        """
        The is_sharepoint_folder method is a wrapper that delegates the
        task of checking if a folder exists in SharePoint to another
        function or method named is_sharepoint_folder. The SharePoint
        client context (self.ctx), the relative path of the folder
        (relative_path), and the folder name (folder_name) are passed
        as arguments to this function. The method then returns the
        boolean result obtained from the is_sharepoint_folder function.

        Args:
            relative_path (str): Relative path where the new folder will be
            created.

            folder_name (str): Name of the new folder to be created.

        Returns:
            bool: True if the folder exists, otherwise return false.
        """
        return is_sharepoint_folder(self.ctx, relative_path, folder_name)

    def upload_file_sharepoint(
        self,
        relative_path: str,
        file_upload: Union[str, io.BytesIO],
        file_name: str = None,
    ) -> None:
        """
        The upload_file_sharepoint method is a wrapper that delegates
        the task of uploading a file to SharePoint to another function
        or method named upload_file_sharepoint. The SharePoint client
        context (self.ctx), the relative path of the folder
        (relative_path), the file to upload (file_upload), and the
        desired file name (file_name) are passed as arguments to this
        function.

        Args:
            relative_path (str): Relative path of the target folder where
            the file will be uploaded.

            file_upload (Union[str, io.BytesIO]): The file to be uploaded.
            It can be either a file path (str) or a BytesIO object.

            file_name (str, optional): Desired name of the file in
            SharePoint. It has a default value of None.
        """
        upload_file_sharepoint(self.ctx, relative_path, file_upload, file_name)

    def create_list(self, list_name: str) -> None:
        """
        The create_list method is a wrapper that delegates the task of
        creating a SharePoint list to another function or method named
        create_list. The SharePoint client context (self.ctx) and the
        name of the list (list_name) are passed as arguments to this
        function.

        Args:
            list_name (str): Name of the sharepoint list.
        """
        create_list(self.ctx, list_name)
    
    def delete_item_list(self, item: ListItem) -> None:
        """
        Deletes a specific SharePoint item.
    
        Args:
            item (ListItem): Item to be deleted.
        """
        delete_item_list(self.ctx, item)

    def delete_folder_with_contents(self, relative_path: str, folder_name: str) -> None:
        """
        Delete a SharePoint folder and its contents recursively.

        Args:
            relative_path (str): Server-relative path of the parent folder.
            folder_name (str): Name of the folder to delete.
        """
        full_path = f"{relative_path}/{folder_name}"
        delete_folder_with_contents(self.ctx, full_path)
