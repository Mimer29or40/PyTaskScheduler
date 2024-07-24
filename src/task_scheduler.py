from __future__ import annotations

import functools
import re
from abc import ABC
from datetime import datetime
from enum import Enum
from enum import Flag
from typing import ClassVar
from typing import Iterable
from typing import Iterator
from typing import List
from typing import Optional
from typing import Sequence
from typing import Sized
from typing import Type

import win32com.client
from dateutil.relativedelta import relativedelta
from pywintypes import com_error  # noqa

# ---------- PYTHON CLASSES ---------- #


class WrapperClass(ABC):
    def __init__(self, obj):
        self._obj = obj

    def __str__(self) -> str:
        return f"{self.__class__.__name__}({self._obj.__str__()})"

    def __repr__(self) -> str:
        return f"{self.__class__.__name__}({self._obj.__repr__()})"


# ---------- CLASSES ---------- #


class TaskService(WrapperClass):
    """For scripting, provides access to the Task Scheduler service for managing registered tasks.

    The TaskService.Connect method should be called before calling any of the other TaskService
    methods.
    """

    _instance: ClassVar[TaskService] = None

    def __new__(cls):
        if cls._instance is None:
            cls._instance = super().__new__(cls)
        return cls._instance

    def __init__(self):
        super().__init__(win32com.client.Dispatch("Schedule.Service"))

    @property
    def connected(self) -> bool:
        """For scripting, gets a Boolean value that indicates if you are connected to the Task
        Scheduler service.

        This property is read-only.

        Property
        A Boolean value that indicates if you are connected to the Task Scheduler service.
        """
        return self._obj.Connected

    @property
    def connected_domain(self) -> str:
        """For scripting, gets the name of the domain to which the TargetServer computer is
        connected.

        Property
        The domain to which the TargetServer computer is connected.
        """
        return self._obj.ConnectedDomain

    @property
    def connected_user(self) -> str:
        """For scripting, gets the name of the user that is connected to the Task Scheduler service.

        Property
        The name of the user that is connected to the Task Scheduler service.
        """
        return self._obj.ConnectedUser

    @property
    def highest_version(self) -> int:
        """For scripting, indicates the highest version of Task Scheduler that a computer supports.

        Property
        The highest version of Task Scheduler that a computer supports. The highest version is a
        value that is split into MajorVersion/MinorVersion on the 16-bit boundary. The Task
        Scheduler service returns 1 for the major version and 2 for the minor version. Use the CLng
        function to get the integer value of the property.
        """
        return self._obj.HighestVersion

    @property
    def target_server(self) -> str:
        """For scripting, gets the name of the computer that is running the Task Scheduler service
        that the user is connected to.

        Property
        The name of the computer that is running the Task Scheduler service that the user is
        connected to.

        Remarks
        This property returns an empty string when the user passes an IP address, Localhost, or "."
        as a parameter, and it returns the name of the computer that is running the Task Scheduler
        service when the user does not pass any parameter value.
        """
        return self._obj.TargetServer

    def connect(
        self, server_name: str = None, user: str = None, domain: str = None, password: str = None
    ) -> None:
        """For scripting, connects to a remote machine and associates all subsequent calls on this
        interface with a remote session. If the serverName parameter is empty, then this method
        will execute on the local computer. If the userId is not specified, then the current token
        is used.

        Remarks
        The TaskService.Connect method should be called before calling any of the other TaskService
        methods.

        If the Connect method fails, you can collect the error identifier to find the meaning of
        the error. The following table lists the error identifiers and their descriptions.

         Error Identifier | Description
        ------------------|-------------------------------------------------------------------------
            0x80070005    | Access is denied to connect to the Task Scheduler service.
        ------------------|-------------------------------------------------------------------------
            0x80041315    | The Task Scheduler service is not running.
        ------------------|-------------------------------------------------------------------------
            0x8007000e    | The application does not have enough memory to complete the operation
                          | or the user, password, or domain has at least one null and one non-null
                          | value.
        ------------------|-------------------------------------------------------------------------
                53        | This error is returned in the following situations:
                          | * The computer name specified in the serverName parameter does not
                          |   exist.
                          | * When you are trying to connect to a Windows Server 2003 or Windows XP
                          |   computer, and the remote computer does not have the File and Printer
                          |   Sharing firewall exception enabled or the Remote Registry service is
                          |   not running.
                          | * When you are trying to connect to a Windows Vista computer, and the
                          |   remote computer does not have the Remote Scheduled Tasks Management
                          |   firewall exception enabled and the File and Printer Sharing firewall
                          |   exception enabled, or the Remote Registry service is not running.
        ------------------|-------------------------------------------------------------------------
                50        | The user, password, or domain parameters cannot be specified when
                          | connecting to a remote Windows XP or Windows Server 2003 computer from
                          | a Windows Vista computer.

        If you are too connecting to a remote Windows Vista computer from a Windows Vista, you need
        to allow the Remote Scheduled Tasks Management firewall exception on the remote computer.
        To allow this exception click Start, Control Panel, Security, Allow a program through
        Windows Firewall, and then select the Remote Scheduled Tasks Management check box. Then
        click the Ok button in the Windows Firewall Settings dialog box.

        If you are connecting to a remote Windows XP or Windows Server 2003 computer from a Windows
        Vista computer, you need to allow the File and Printer Sharing firewall exception on the
        remote computer. To allow this exception click Start, Control Panel, double-click Windows
        Firewall, select the Exceptions tab, and then select the File and Printer Sharing firewall
        exception. Then click the OK button in the Windows Firewall dialog box. The Remote Registry
        service must also be running on the remote computer.

        :param server_name: The name of the computer that you want to connect to. If the serverName
          parameter is empty, then this method will execute on the local computer.
        :param user: The username that is used during the connection to the computer. If the user
          is not specified, then the current token is used.
        :param domain: The domain of the user specified in the user parameter.
        :param password: The password that is used to connect to the computer. If the username and
          password are not specified, then the current token is used.
        """
        self._obj.Connect(server_name, user, domain, password)

    def get_folder(self, path: str) -> TaskFolder:
        """For scripting, gets a folder of registered tasks.

        :param path: The path to the folder to be retrieved. Do not use a backslash following the
          last folder name in the path. The root task folder is specified with a backslash (\\). An
          example of a task folder path, under the root task folder, is \\MyTaskFolder. The "."
          character cannot be used to specify the current task folder and the ".." characters
          cannot be used to specify the parent task folder in the path.
        :returns: A TaskFolder object for the requested folder.
        """
        try:
            return TaskFolder(self._obj.GetFolder(path))
        except com_error:
            raise TaskFolderNotFound(f"Folder does not exist: {path!r}") from None

    def get_running_tasks(self, hidden: bool = False) -> RunningTaskCollection:
        """For scripting, gets a collection of running tasks.

        Note

        TaskService.GetRunningTasks will only return a collection of running tasks that are running
        at or below a user's security context. For example, for members of the Administrators
        group, GetRunningTasks will return a collection of all running tasks, but for members of
        the Users group, GetRunningTasks will only return a collection of tasks that are running
        under the Users group security context.

        :param hidden: Pass in True to return all running tasks, including hidden tasks. Pass in
          False to return a collection of running tasks that are not hidden tasks.
        :returns: A RunningTaskCollection object that contains the currently running tasks.
        """
        return RunningTaskCollection(self._obj.GetRunningTasks(1 if hidden else 0))

    def new_task(self, flags: int = 0) -> TaskDefinition:  # TODO - Create Flag Class
        """For scripting, returns an empty task definition object to be filled in with settings and
        properties and then registered using the TaskFolder.RegisterTaskDefinition method.

        :param flags: This parameter is reserved for future use and must be set to 0.
        :returns: The task definition that specifies all the information required to create a new
          task.
        """
        return TaskDefinition(self._obj.NewTask(flags))


class TaskFolder(WrapperClass):
    """Scripting object that provides the methods that are used to register (create) tasks in the
    folder, remove tasks from the folder, and create or remove sub-folders from the folder.
    """

    def __eq__(self, other: TaskFolder) -> bool:
        if not isinstance(other, TaskFolder):
            return False
        return self.path == other.path

    @property
    def name(self) -> str:
        """For scripting, gets the name that is used to identify the folder that
        contains a task.

        :returns: The name that is used to identify the folder.
        """
        return self._obj.Name

    @property
    def path(self) -> str:
        """For scripting, gets the path to where the folder is stored.

        :returns: The path to where the folder is stored. The root task folder is specified with a
          backslash (\\). An example of a task folder path, under the root task folder, is
          \\MyTaskFolder.
        """
        return self._obj.Path

    def create_folder(self, folder_name: str, security_descriptor: str = None) -> TaskFolder:
        """For scripting, creates a folder for related tasks.

        :param folder_name: The name that is used to identify the folder. If
          "FolderName\\SubFolder1\\SubFolder2" is specified, the entire folder tree will be created
          if the folders do not exist. This parameter can be a relative path to the current
          TaskFolder instance. The root task folder is specified with a backslash (\\). An example
          of a task folder path, under the root task folder, is \\MyTaskFolder. The "." character
          cannot be used to specify the current task folder and the ".." characters cannot be used
          to specify the parent task folder in the path.
        :param security_descriptor: The security descriptor that is associated with the folder.
        :returns: A TaskFolder object that represents the new sub-folder.
        """
        try:
            return TaskFolder(self._obj.CreateFolder(folder_name, security_descriptor))
        except com_error:
            raise TaskFolderExists(f"Folder already exists: {folder_name!r}") from None

    def delete_folder(self, folder_name: str, flags: int = 0) -> None:
        """For scripting, deletes a sub-folder from the parent folder.

        :param folder_name: The name of the sub-folder to be removed. The root task folder is
          specified with a backslash (\\). This parameter can be a relative path to the folder you
          want to delete. An example of a task folder path, under the root task folder, is
          \\MyTaskFolder. The "." character cannot be used to specify the current task folder and
          the ".." characters cannot be used to specify the parent task folder in the path
        :param flags: Not supported.
        """
        try:
            self._obj.DeleteFolder(folder_name, flags)
        except com_error:
            raise TaskFolderNotFound(f"Folder not found: {folder_name!r}") from None

    def delete_task(self, name: str, flags: int = 0) -> None:
        """For scripting, deletes a task from the folder.

        :param name: The name of the task that is specified when the task was registered. The "."
          character cannot be used to specify the current task folder and the ".." characters
          cannot be used to specify the parent task folder in the path.
        :param flags: Not supported. Value is 0
        """
        self._obj.DeleteTask(name, flags)

    def get_folder(self, path: str) -> TaskFolder:
        """For scripting, gets a folder that contains tasks at a specified location.

        param path: The path (location) to the folder. Do not use a backslash following the last
          folder name in the path. The root task folder is specified with a backslash (\\). An
          example of a task folder path, under the root task folder, is \\MyTaskFolder. The "."
          character cannot be used to specify the current task folder and the ".." characters
          cannot be used to specify the parent task folder in the path.
        :return: The folder at the specified location. The folder is a Folder object.
        """
        try:
            return TaskFolder(self._obj.GetFolder(path))
        except com_error:
            raise TaskFolderNotFound(f"Folder not found: {path!r}") from None

    def get_folders(self, flags: int = 0) -> TaskFolderCollection:
        """For scripting, gets all the sub-folders in the folder.

        :param flags: TODO
        """
        return TaskFolderCollection(self._obj.GetFolders(flags))

    def get_security_descriptor(self, security_info: int) -> str:
        """For scripting, gets the security descriptor for the folder.

        :param security_info: TODO
        """
        return self._obj.GetSecurityDescriptor(security_info)

    def get_task(self, path: str) -> RegisteredTask:
        """For scripting, gets a task at a specified location in a folder.

        :param path: The path (location) to the task in a folder. The root task folder is specified
          with a backslash (\\). An example of a task folder path, under the root task folder, is
          \\MyTaskFolder. The "." character cannot be used to specify the current task folder and
          the ".." characters cannot be used to specify the parent task folder in the path.
        :return: The task at the specified location. The task is a RegisteredTask object.
        """
        return RegisteredTask(self._obj.GetTask(path))

    def get_tasks(self, flags: int = 0) -> RegisteredTaskCollection:
        """For scripting, gets all the tasks in the folder.

        :param flags: TODO
        """
        return RegisteredTaskCollection(self._obj.GetTasks(flags))

    def register_task(
        self,
        path: str,
        xml_text: str,
        flags: Creation,
        user_id: str,
        password: str,
        logon_type: LogonType,
        security_descriptor: str = None,
    ) -> RegisteredTask:
        """For scripting, registers (creates) a new task in the folder using XML to define the task.

        For a task, that contains a message box action, the message box will be displayed if the
        task is activated and the task has an interactive logon type. To set the task logon type to
        interactive, specify 3 (TASK_LOGON_INTERACTIVE_TOKEN) or 4 (TASK_LOGON_GROUP) in the
        LogonType property of the task principal, or in the logonType parameter of
        TaskFolder.RegisterTask or TaskFolder.RegisterTaskDefinition.

        Only a member of the Administrators group can create a task with a boot trigger.

        You can successfully register a task with a group specified in the userId parameter and 3
        (TASK_LOGON_INTERACTIVE_TOKEN) specified in the logonType parameter of
        TaskFolder.RegisterTask or TaskFolder.RegisterTaskDefinition, but the task will not run.

        :param path: The name of the task. If this value is Nothing, the task will be registered in
          the root task folder and the task name will be a GUID value that is created by the Task
          Scheduler service. A task name cannot begin or end with a space character. The "."
          character cannot be used to specify the current task folder and the ".." characters
          cannot be used to specify the parent task folder in the path.
        :param xml_text: An XML-formatted description of the task.
          The following topics contain tasks defined using XML.
          * Time Trigger Example (XML)
          * Event Trigger Example (XML)
          * Daily Trigger Example (XML)
          * Registration Trigger Example (XML)
          * Weekly Trigger Example (XML)
          * Logon Trigger Example (XML)
          * Boot Trigger Example (XML)
        :param flags: A Creation constant.
        :param user_id: The user credentials that are used to register the task. If the task is
          defined as a Task Scheduler 1.0 task, then do not use a group name (rather than a
          specific username) in this userId parameter. A task is defined as a Task Scheduler 1.0
          task when the version attribute of the Task element in the task's XML is set to 1.1.
        :param password: The password for the userId that is used to register the task. When the
          TASK_LOGON_SERVICE_ACCOUNT logon type is used, the password must be an empty VARIANT
          value such as VT_NULL or VT_EMPTY.
        :param logon_type: Defines what logon technique is used to run the registered task.
        :param security_descriptor: The security descriptor that is associated with the registered
          task. You can specify the access control list (ACL) in the security descriptor for a task
          in order to allow or deny certain users and groups access to a task. Note: If the Local
          System account is denied access to a task, then the Task Scheduler service can produce
          unexpected results.
        :return: A RegisteredTask object that represents the new task.
        """
        return RegisteredTask(
            self._obj.RegisterTask(
                path,
                xml_text,
                flags.value,
                user_id,
                password,
                logon_type,
                security_descriptor,
            )
        )

    def register_task_definition(
        self,
        path: str,
        definition: TaskDefinition,
        flags: Creation,
        user_id: str,
        password: str,
        logon_type: LogonType,
        security_descriptor: str = None,
    ) -> RegisteredTask:
        """For scripting, registers (creates) a task in a specified location using the
        TaskDefinition object to define a task.

        For a task, that contains a message box action, the message box will be displayed if the
        task is activated and the task has an interactive logon type. To set the task logon type to
        interactive, specify 3 (TASK_LOGON_INTERACTIVE_TOKEN) or 4 (TASK_LOGON_GROUP) in the
        LogonType property of the task principal, or in the logonType parameter of
        TaskFolder.RegisterTask or TaskFolder.RegisterTaskDefinition.

        Only a member of the Administrators group can create a task with a boot trigger.

        You can successfully register a task with a group specified in the userId parameter and 3
        (TASK_LOGON_INTERACTIVE_TOKEN) specified in the logonType parameter of
        TaskFolder.RegisterTask or TaskFolder.RegisterTaskDefinition, but the task will not run.

        :param path: The name of the task. If this value is Nothing, the task will be registered in
          the root task folder and the task name will be a GUID value that is created by the Task
          Scheduler service. A task name cannot begin or end with a space character. The "."
          character cannot be used to specify the current task folder and the ".." characters
          cannot be used to specify the parent task folder in the path.
        :param definition: The definition of the task that is registered.
        :param flags: A Creation constant.
        :param user_id: The user credentials that are used to register the task. If present, these
          credentials take priority over the credentials specified in the task definition object
          pointed to by the definition parameter. Note: If the task is defined as a Task Scheduler
          1.0 task, then do not use a group name (rather than a specific username) in this userId
          parameter. A task is defined as a Task Scheduler 1.0 task when the Compatibility property
          is set to 1 in the task's settings.
        :param password: The password for the userId that is used to register the task. When the
          TASK_LOGON_SERVICE_ACCOUNT logon type is used, the password must be an empty VARIANT
          value such as VT_NULL or VT_EMPTY.
        :param logon_type: Defines what logon technique is used to run the registered task.
        :param security_descriptor: The security descriptor that is associated with the registered
          task. You can specify the access control list (ACL) in the security descriptor for a task
          in order to allow or deny certain users and groups access to a task. Note: If the Local
          System account is denied access to a task, then the Task Scheduler service can produce
          unexpected results.
        :return: A RegisteredTask object that represents the new task.
        """
        return RegisteredTask(
            self._obj.RegisterTaskDefinition(
                path,
                definition._obj,
                flags.value,
                user_id,
                password,
                logon_type.value,
                security_descriptor,
            )
        )

    def set_security_description(self, security_descriptor: str, flags: int) -> None:
        """For scripting, sets the security descriptor for the folder.

        :param security_descriptor: TODO
        :param flags: TODO
        """
        self._obj.SetSecurityDescriptor(security_descriptor, flags)


class TaskFolderCollection(WrapperClass, Iterable, Sized):
    """Scripting object that provides information and control for a collection of folders that
    contain tasks.
    """

    def __len__(self) -> int:
        return self.count

    def __getitem__(self, index: int) -> TaskFolder:
        return self.item(index)

    def __iter__(self) -> Iterator[TaskFolder]:
        return iter(map(TaskFolder, self._obj))

    @property
    def count(self) -> int:
        """For scripting, gets the number of folders in the collection.

        :returns: The number of triggers in the collection.
        """
        return self._obj.Count

    def item(self, index: int) -> TaskFolder:
        """For scripting, gets the specified folder from the collection.

        Remarks

        Collections are 1-based. In other words, the index for the first item in the collection is
        1.

        :returns: A TaskFolder object that represents the requested folder.
        """
        try:
            return TaskFolder(self._obj.Item(index))
        except com_error:
            raise IndexError(f"No Folder at index: {index}") from None


class TaskDefinition(WrapperClass):
    """Scripting object that defines all the components of a task, such as the task settings,
    triggers, actions, and registration information.

    Remarks
    When reading or writing your own XML for a task, a task definition is specified using the Task
    element of the Task Scheduler schema.
    """

    @functools.cached_property
    def actions(self) -> ActionCollection:
        """For scripting, gets or sets a collection of actions that are performed by the task.

        :returns: A collection of actions that are performed by the task.
        """
        return ActionCollection(self._obj.Actions)

    @property
    def data(self) -> str:
        """For scripting, gets or sets the data that is associated with the task. This data is
        ignored by the Task Scheduler service, but is used by third-parties who wish to extend the
        task format.

        :returns: The data that is associated with the task.
        """
        return self._obj.Data

    @data.setter
    def data(self, value: str) -> None:
        self._obj.Data = value

    @functools.cached_property
    def principal(self) -> Principal:
        """For scripting, gets or sets the principal for the task that provides the security
        credentials for the task.

        :returns: The principal for the task that provides the security credentials for the task.
        """
        return Principal(self._obj.Principal)

    @functools.cached_property
    def registration_info(self) -> RegistrationInfo:
        """For scripting, gets or sets the registration information that is used to describe a task,
        such as the description of the task, the author of the task, and the date the task is
        registered.

        :returns: The registration information that is used to describe a task, such as the
         description of the task, the author of the task, and the date the task is registered.
        """
        return RegistrationInfo(self._obj.RegistrationInfo)

    @functools.cached_property
    def settings(self) -> TaskSettings:
        """For scripting, gets or sets the settings that define how the Task Scheduler service
        performs the task.

        :returns: The settings that define how the Task Scheduler service performs the task.
        """
        return TaskSettings(self._obj.Settings)

    @functools.cached_property
    def triggers(self) -> TriggerCollection:
        """For scripting, gets or sets a collection of triggers that are used to start a task.

        :returns: A collection of triggers that are used to start a task.
        """
        return TriggerCollection(self._obj.Triggers)

    @property
    def xml_text(self) -> str:
        """For scripting, gets or sets the XML-formatted definition of the task.

        Remarks

        The XML for a task is defined by the Task Scheduler Schema. For an example of task XML, see
        Daily Trigger Example (XML).

        :returns: The XML-formatted definition of the task.
        """
        return self._obj.XmlText

    @xml_text.setter
    def xml_text(self, value: str) -> None:
        self._obj.XmlText = value


class RunningTask(WrapperClass):
    """Scripting object that provides the methods to get information from and control a running
    task.
    """

    @property
    def current_action(self) -> str:
        """For scripting, gets the name of the current action that the running task is performing.

        :returns: The name of the current action that the running task is performing.
        """
        return self._obj.CurrentAction

    @property
    def engine_pid(self) -> int:
        """For scripting, gets the process ID for the engine (process) which is running the task.

        :returns: The process ID for the engine which is running the task.
        """
        return self._obj.EnginePID

    @property
    def instance_guid(self) -> str:
        """For scripting, gets the GUID identifier for this instance of the task.

        :returns: The GUID identifier for this instance of the task. An identifier is generated by
         the Task Scheduler service each time the task is run.
        """
        return self._obj.InstanceGuid

    @property
    def name(self) -> str:
        """For scripting, gets the name of the task.

        :returns: The name of the task.
        """
        return self._obj.Name

    @property
    def path(self) -> str:
        """For scripting, gets the path to where the task is stored.

        :returns: The path to where the task is stored.
        """
        return self._obj.Path

    @property
    def state(self) -> State:
        """For scripting, gets an identifier for the state of the running task.

        :returns: An identifier for the state of the running task.
        """
        return State(self._obj.State)

    def refresh(self) -> None:
        """For scripting, refreshes all the local instance variables of the task."""
        self._obj.Refresh()

    def stop(self) -> None:
        """For scripting, stops this instance of the task."""
        self._obj.Stop()


class RunningTaskCollection(WrapperClass, Iterable, Sized):
    """Scripting object that provides a collection that is used to control running tasks."""

    def __len__(self) -> int:
        return self.count

    def __getitem__(self, index: int) -> RunningTask:
        return self.item(index)

    def __iter__(self) -> Iterator[RunningTask]:
        return iter(map(RunningTask, self._obj))

    @property
    def count(self) -> int:
        """For scripting, gets the number of running tasks in the collection.

        :returns: The number of running tasks in the collection.
        """
        return self._obj.Count

    def item(self, index: int) -> RunningTask:
        """For scripting, gets the specified running task from the collection.

        Remarks

        Collections are 1-based. In other words, the index for the first item in the collection is
        1.

        :returns: A RunningTask object that contains the running task.
        """
        return RunningTask(self._obj.Item(index))


class RegisteredTask(WrapperClass):
    """Scripting object that provides the methods that are used to run the task immediately, get
    any running instances of the task, get or set the credentials that are used to register the
    task, and the properties that describe the task.
    """

    @functools.cached_property
    def definition(self) -> TaskDefinition:
        """For scripting, gets the definition of the task.

        :returns: The definition of the task.
        """
        return TaskDefinition(self._obj.Definition)

    @property
    def enabled(self) -> bool:
        """For scripting, gets or sets a Boolean value that indicates if the registered task is
        enabled.

        :returns: A Boolean value that indicates if the registered task is enabled.
        """
        return self._obj.Enabled

    @enabled.setter
    def enabled(self, value: bool) -> None:
        self._obj.Enabled = value

    @property
    def last_run_time(self) -> datetime:
        """For scripting, gets the time the registered task was last run.

        :returns: The time the registered task was last run.
        """
        return self._obj.LastRunTime

    @property
    def last_task_result(self) -> int:
        """For scripting, gets the results that were returned the last time the registered task was
        run.

        :returns: The results that were returned the last time the registered task was run.
        """
        return self._obj.LastTaskResult

    @property
    def name(self) -> str:
        """For scripting, gets the name of the registered task.

        :returns: The name of the registered task.
        """
        return self._obj.Name

    @property
    def next_run_time(self) -> datetime:
        """For scripting, gets the time when the registered task is next scheduled to run.

        Remarks

        If the registered task contains triggers that are individually disabled, these triggers
        will still affect the next scheduled run time that is returned even though they are
        disabled.

        :returns: The time when the registered task is next scheduled to run.
        """
        return self._obj.NextRunTime

    @property
    def number_of_missed_runs(self) -> int:
        """For scripting, gets the number of times the registered task has missed a scheduled run.

        :returns: The number of times the registered task missed a scheduled run.
        """
        return self._obj.NumberOfMissedRuns

    @property
    def path(self) -> str:
        """For scripting, gets the path to where the registered task is stored.

        :returns: The path to where the registered task is stored.
        """
        return self._obj.Path

    @property
    def state(self) -> State:
        """For scripting, gets the operational state of the registered task.

        :returns: A TASK_STATE constant that defines the operational state of the task.
        """
        return State(self._obj.State)

    @property
    def xml(self) -> str:
        """For scripting, gets the XML-formatted registration information for the registered task.

        :returns: The XML-formatted registration information for the registered task.
        """
        return self._obj.XML

    def get_instances(self, flags: int = 0) -> RunningTaskCollection:
        """For scripting, returns all currently running instances of the registered task.

        Note

        RegisteredTask.GetInstances will only return instances of the currently running registered
        task that are running at or below a user's security context. For example, for members of
        the Administrators group, GetInstances will return all instances of the currently running
        registered task, but for members of the Users group, GetInstances will only return
        instances of the currently running registered task that are running under the Users group
        security context.

        :param flags: This parameter is reserved for future use and must be set to 0.
        :returns: A RunningTaskCollection object that contains all currently running instances of
          the task.
        """
        return RunningTaskCollection(self._obj.GetInstances(flags))

    def get_run_times(self, pst_start: datetime, pst_end: datetime):  # TODO - Types
        """Gets the times that the registered task is scheduled to run during a specified time.

        Remarks

        If the registered task contains triggers that are individually disabled, these triggers
        will still affect the next scheduled run time that is returned even though they are
        disabled.

        :param pst_start: The starting time for the query.
        :param pst_end: The ending time for the query.
        :returns: The requested number of runs on input and the returned number of runs on output.
          The scheduled times that the task will run. A NULL LPSYSTEMTIME object should be
          passed into this parameter. On return, this array contains pCount run times. You must
          free this array by a calling the CoTaskMemFree function.
          If the method succeeds, it returns S_OK. If the method returns S_FALSE, the
          pRunTimes parameter contains pCount items, but there were more runs of the task, that
          were not returned. Otherwise, it returns an HRESULT error code.
        """
        # return self._obj.GetRunTimes(to_date_str(pst_start), to_date_str(pst_end))
        return self._obj.GetRunTimes(pst_start, pst_end)

    def get_security_descriptor(self, security_info: SecurityInformation) -> str:
        """For scripting, gets the security descriptor that is used as credentials for the
        registered task.

        :param security_info: The security information from SECURITY_INFORMATION.
        :returns: The security descriptor for the registered task.
        """
        return self._obj.GetSecurityDescriptor(security_info.value)

    def run(self, params: Optional[str]) -> RunningTask:
        """For scripting, runs the registered task immediately.

        Remarks

        The RegisteredTask.Run function is equivalent to the RegisteredTask.RunEx function with the
        flags parameter equal to 0 and nothing specified for the user parameter.

        This method will return without error, but the task will not run if the
        TaskSettings.AllowDemandStart property is set to false for the registered task.

        :param params: The parameters used as values in the task actions. To not specify any
          parameter values for the task actions, set this parameter to Nothing. Otherwise, a single
          string value or an array of string values can be specified.
          The string values that you specify are paired with names and stored as name-value pairs.
          If you specify a single string value, then Arg0 will be the name assigned to the value.
          The value can be used in the task action where the $(Arg0) variable is used in the action
          properties.
          If you pass in values such as "0", "100", and "250" as an array of string values, then
          "0" will replace the $(Arg0) variables, "100" will replace the $(Arg1) variables, and
          "250" will replace the $(Arg2) variables used in the action properties.
          A maximum of 32 string values can be specified.
          For more information and a list of action properties that can use $(Arg0), $(Arg1), ...,
          $(Arg32) variables in their values, see Task Actions.
        :returns: A RunningTask object that defines the new instance of the task.
        """
        return RunningTask(self._obj.Run(params))

    def run_ex(self, params: Optional[str], flags: RunFlags, session_id: int) -> RunningTask:
        """For scripting, runs the registered task immediately using specified flags and a session
        identifier.

        Remarks

        This method will return without error, but the task will not run if the
        TaskSettings.AllowDemandStart property is set to false for the registered task.

        :param params: The parameters used as values in the task actions. To not specify any
          parameter values for the task actions, set this parameter to Nothing. Otherwise, a single
          string value or an array of string values can be specified.
          The string values that you specify are paired with names and stored as name-value pairs.
          If you specify a single string value, then Arg0 will be the name assigned to the value.
          The value can be used in the task action where the $(Arg0) variable is used in the action
          properties.
          If you pass in values such as "0", "100", and "250" as an array of string values, then
          "0" will replace the $(Arg0) variables, "100" will replace the $(Arg1) variables, and
          "250" will replace the $(Arg2) variables used in the action properties.
          A maximum of 32 string values can be specified.
          For more information and a list of action properties that can use $(Arg0), $(Arg1), ...,
          $(Arg32) variables in their values, see Task Actions.
        :param flags: A RunFlags constant that defines how the task is run.
        :param session_id: The terminal server session in which you want to launch the task.
          If the TASK_RUN_USE_SESSION_ID constant (0x4) is not passed into "flags", then
          the value specified in this parameter is ignored. If the TASK_RUN_USE_SESSION_ID constant
          is passed into the flags parameter and the sessionID value is less than or equal to 0,
          then an invalid argument error will be returned.
          If the TASK_RUN_USE_SESSION_ID constant is passed into the flags parameter and the
          sessionID value is a valid session ID greater than 0 and if no value is specified for the
          user parameter, then the Task Scheduler service will try to launch the task interactively
          as the user who is logged on to the specified session.
          If the TASK_RUN_USE_SESSION_ID constant is passed into the flags parameter and the
          sessionID value is a valid session ID greater than 0 and if a user is specified in the
          user parameter, then the Task Scheduler service will try to launch the task interactively
          as the user who is specified in the user parameter.
        :returns: A RunningTask object that defines the new instance of the task.
        """
        return RunningTask(self._obj.RunEx(params, flags.value, session_id))

    def set_security_descriptor(self, security_descriptor: str, flags: Creation) -> None:
        """For scripting, sets the security descriptor that is used as credentials for the
        registered task.

        Remarks

        You can specify the access control list (ACL) in the security descriptor for a task in
        order to allow or deny certain users and groups access to a task.

        :param security_descriptor: The security descriptor that is used as credentials for the
          registered task.
          Note: If the Local System account is denied access to a task, then the Task Scheduler
          service can produce unexpected results.
        :param flags: Flags that specify how to set the security descriptor. The
          TASK_DONT_ADD_PRINCIPAL_ACE flag (0x10) from the TASK_CREATION enumeration can be
          specified.
        """
        return self._obj.SetSecurityDescriptor(security_descriptor, flags.value)

    def stop(self, flags: int = 0) -> None:
        """For scripting, stops the registered task immediately.

        Remarks

        The RegisteredTask.Stop function stops all instances of the task.

        System account users can stop a task, users with Administrator group privileges can stop a
        task, and if a user has rights to execute and read a task, then the user can stop the task.
        A user can stop the task instances that are running under the same credentials as the user
        account. In all other cases, the user is denied access to stop the task.

        :param flags: Reserved. Must be zero.
        """
        return self._obj.Stop(flags)


class RegisteredTaskCollection(WrapperClass, Iterable, Sized):
    """Scripting object that contains all the tasks that are registered."""

    def __len__(self) -> int:
        return self.count

    def __getitem__(self, index: int) -> RegisteredTask:
        return self.item(index)

    def __iter__(self) -> Iterator[RegisteredTask]:
        return iter(map(RegisteredTask, self._obj))

    @property
    def count(self) -> int:
        """For scripting, gets the number of registered tasks in the collection.

        :returns: The number of registered tasks in the collection.
        """
        return self._obj.Count

    def item(self, index: int) -> RegisteredTask:
        """For scripting, gets the specified registered task from the collection.

        Remarks

        Collections are 1-based. In other words, the index for the first item in the collection is
        1.

        :returns: A RegisteredTask object that contains the registered task.
        """
        try:
            return RegisteredTask(self._obj.Item(index))
        except com_error:
            raise IndexError(f"No RegisteredTask at index: {index}") from None


class TaskVariables(WrapperClass):
    """Scripting object that defines task variables that can be passed as parameters to task
    handlers and external executables that are launched by tasks.
    """

    def get_context(self):  # TODO - Return Type
        """For scripting, used to share the context between different steps and tasks that are in
        the same job instance. This method is not implemented.

        :returns: The context that is used to share the context between different steps and tasks
          that are in the same job instance.
        """
        return self._obj.GetContext()

    def get_input(self):  # TODO - Return Type
        """For scripting, gets the input variables for a task. This method is not implemented.

        :returns: The input variables for a task.
        """
        return self._obj.GetInput()

    def set_input(self, input):  # TODO - Input Type
        """For scripting, sets the output variables for a task. This method is not implemented.

        :param input: The output variables for a task.
        """
        return self._obj.SetInput(input)


class RegistrationInfo(WrapperClass):
    """Scripting object that provides the administrative information that can be used to describe
    the task. This information includes details such as a description of the task, the author of
    the task, the date the task is registered, and the security descriptor of the task.

    Remarks
    Registration information can be used to identify a task through the Task Scheduler UI, or as
    search criteria when enumerating over the registered tasks.

    When reading or writing XML for a task, registration information for the task is specified in
    the RegistrationInfo element of the Task Scheduler schema.
    """

    @property
    def author(self) -> str:
        """For scripting, gets or sets the author of the task.

        Remarks

        When reading or writing XML for a task, the task author is specified using the Author
        element of the Task Scheduler schema.

        When setting this property value, the value can be text that is retrieved from a resource
        .dll file. A specialized string is used to reference the text from the resource file. The
        format of the string is $(@ [Dll], [ResourceID]) where [Dll] is the path to the .dll file
        that contains the resource and [ResourceID] is the identifier for the resource text. For
        example, the setting this property value to $(@ %SystemRoot%\\System32\\ResourceName.dll,
        -101) will set the property to the value of the resource text with an identifier equal to
        -101 in the %SystemRoot%\\System32\\ResourceName.dll file.

        :returns: The author of the task.
        """
        return self._obj.Author

    @author.setter
    def author(self, value: str) -> None:
        self._obj.Author = value

    @property
    def date(self) -> Optional[datetime]:
        """For scripting, gets or sets the date and time when the task is registered.

        Remarks

        When reading or writing XML for a task, the registration date is specified using the Date
        element of the Task Scheduler schema.

        :returns: The registration date of the task.
        """
        return from_date_str(self._obj.Date)

    @date.setter
    def date(self, value: Optional[datetime]) -> None:
        self._obj.Date = to_date_str(value)

    @property
    def description(self) -> str:
        """For scripting, gets or sets the description of the task.

        Remarks

        When reading or writing XML for a task, the description of the task is specified using the
        Description element of the Task Scheduler schema.

        When setting this property value, the value can be text that is retrieved from a resource
        .dll file. A specialized string is used to reference the text from the resource file. The
        format of the string is $(@ [Dll], [ResourceID]) where [Dll] is the path to the .dll file
        that contains the resource and [ResourceID] is the identifier for the resource text. For
        example, the setting this property value to $(@ %SystemRoot%\\System32\\ResourceName.dll,
        -101) will set the property to the value of the resource text with an identifier equal to
        -101 in the %SystemRoot%\\System32\\ResourceName.dll file.

        :returns: The description of the task.
        """
        return self._obj.Description

    @description.setter
    def description(self, value: str) -> None:
        self._obj.Description = value

    @property
    def documentation(self) -> str:
        """For scripting, gets or sets any additional documentation for the task.

        Remarks

        When reading or writing XML for a task, the additional documentation for the task is
        specified using the Documentation element of the Task Scheduler schema.

        When setting this property value, the value can be text that is retrieved from a resource
        .dll file. A specialized string is used to reference the text from the resource file. The
        format of the string is $(@ [Dll], [ResourceID]) where [Dll] is the path to the .dll file
        that contains the resource and [ResourceID] is the identifier for the resource text. For
        example, the setting this property value to $(@ %SystemRoot%\\System32\\ResourceName.dll,
        -101) will set the property to the value of the resource text with an identifier equal to
        -101 in the %SystemRoot%\\System32\\ResourceName.dll file.

        :returns: Any additional documentation that is associated with the task.
        """
        return self._obj.Documentation

    @documentation.setter
    def documentation(self, value: str) -> None:
        self._obj.Documentation = value

    @property
    def security_descriptor(self) -> Optional[str]:
        """For scripting, gets or sets the security descriptor of the task. If a different security
        descriptor is supplied during task registration, it will supersede the security descriptor
        set with this property.

        Remarks

        When reading or writing XML for a task, the security descriptor of the task is specified
        using the SecurityDescriptor element of the Task Scheduler schema.

        :returns: The security descriptor that is associated with the task.
        """
        return self._obj.SecurityDescriptor

    @security_descriptor.setter
    def security_descriptor(self, value: Optional[str]) -> None:
        self._obj.SecurityDescriptor = value

    @property
    def source(self) -> str:
        """For scripting, gets or sets where the task originated from. For example, a task may
        originate from a component, service, application, or user.

        Remarks

        When reading or writing XML for a task, the task source information is specified using the
        Source element of the Task Scheduler schema.

        When setting this property value, the value can be text that is retrieved from a resource
        .dll file. A specialized string is used to reference the text from the resource file. The
        format of the string is $(@ [Dll], [ResourceID]) where [Dll] is the path to the .dll file
        that contains the resource and [ResourceID] is the identifier for the resource text. For
        example, the setting this property value to $(@ %SystemRoot%\\System32\\ResourceName.dll,
        -101) will set the property to the value of the resource text with an identifier equal to
        -101 in the %SystemRoot%\\System32\\ResourceName.dll file.

        :returns: Where the task originated from. For example, from a component, service,
          application, or user.
        """
        return self._obj.Source

    @source.setter
    def source(self, value: str) -> None:
        self._obj.Source = value

    @property
    def uri(self) -> str:
        """For scripting, gets or sets the URI of the task.

        This property is read/write.

        Remarks

        When reading or writing XML for a task, the task URI is specified using the URI element of
        the Task Scheduler schema.

        :returns: The URI of the task.
        """
        return self._obj.URI

    @uri.setter
    def uri(self, value: str) -> None:
        self._obj.URI = value

    @property
    def version(self) -> str:
        """For scripting, gets or sets the version number of the task.

        Remarks

        When reading or writing XML for a task, the version number of the task is specified using
        the Version element of the Task Scheduler schema.

        :returns: The version number of the task.
        """
        return self._obj.Version

    @version.setter
    def version(self, value: str) -> None:
        self._obj.Version = value

    @property
    def xml_text(self) -> str:
        """For scripting, gets or sets an XML-formatted version of the registration information for
        the task.

        :returns: An XML-formatted version of the task registration information.
        """
        return self._obj.XmlText

    @xml_text.setter
    def xml_text(self, value: str) -> None:
        self._obj.XmlText = value


class RepetitionPattern(WrapperClass):
    """Scripting object that defines how often the task is run and how long the repetition pattern
    is repeated after the task is started.

    If you specify a repetition duration for a task, you must also specify the repetition interval.

    If you register a task that contains a trigger with a repetition interval equal to one minute
    and a repetition duration equal to four minutes, the task will be launched five times. The five
    repetitions can be defined by the following pattern.

    1. A task starts at the beginning of the first minute.
    2. The next task starts at the end of the first minute.
    3. The next task starts at the end of the second minute.
    4. The next task starts at the end of the third minute.
    5. The next task starts at the end of the fourth minute.

    Windows Server 2003, Windows XP and Windows 2000: If you register a task that contains a
    trigger with a repetition interval equal to one minute and a repetition duration equal to four
    minutes, the task will be launched four times.

    When reading or writing XML for a task, the repetition pattern is specified using the
    Repetition element of the Task Scheduler schema.
    """

    @property
    def duration(self) -> relativedelta:
        """For scripting, gets or sets how long the pattern is repeated.

        If you specify a repetition duration for a task, you must also specify the repetition
        interval.

        When reading or writing XML for a task, the repetition duration is specified in the
        Duration element of the Task Scheduler schema.

        :return: How long the pattern is repeated. The format for this string is PnYnMnDTnHnMnS,
          where nY is the number of years, nM is the number of months, nD is the number of days,
          "T" is the date/time separator, nH is the number of hours, nM is the number of minutes,
          and nS is the number of seconds (for example, PT5M specifies 5 minutes and P1M4DT2H5M
          specifies one month, four days, two hours, and five minutes). The minimum time allowed is
          one minute. If no value is specified for the duration, then the pattern is repeated
          indefinitely.
        """
        return from_duration_str(self._obj.Duration)

    @duration.setter
    def duration(self, value: relativedelta) -> None:
        self._obj.Duration = to_duration_str(value)

    @property
    def interval(self) -> relativedelta:
        """For scripting, gets or sets the amount of time between each restart of the task.

        If you specify a repetition duration for a task, you must also specify the repetition
        interval.

        When reading or writing XML for a task, the repetition interval is specified in the
        Interval element of the Task Scheduler schema.

        :return: The amount of time between each restart of the task. The format for this string is
          P<days>DT<hours>H<minutes>M<seconds>S (for example, "PT5M" is 5 minutes, "PT1H" is 1 hour,
          and "PT20M" is 20 minutes). The maximum time allowed is 31 days, and the minimum time
          allowed is 1 minute.
        """
        return from_duration_str(self._obj.Interval)

    @interval.setter
    def interval(self, value: relativedelta) -> None:
        self._obj.Interval = to_duration_str(value)

    @property
    def stop_at_duration_end(self) -> bool:
        """For scripting, gets or sets a Boolean value that indicates if a running instance of the
        task is stopped at the end of the repetition pattern duration.

        When reading or writing XML for a task, this information is specified in the
        StopAtDurationEnd element of the Task Scheduler schema.

        :return: A Boolean value that indicates if a running instance of the task is stopped at the
          end of the repetition pattern duration.
        """
        return self._obj.StopAtDurationEnd

    @stop_at_duration_end.setter
    def stop_at_duration_end(self, value: bool) -> None:
        self._obj.StopAtDurationEnd = value


class Principal(WrapperClass):
    """Scripting object that provides the security credentials for a principal. These security
    credentials define the security context for the tasks that are associated with the principal.

    Remarks

    When specifying an account, remember to properly use the double backslash in code to specify
    the domain and username. For example, use DOMAIN\\UserName to specify a value for the UserId
    property.

    When reading or writing XML for a task, the security credentials for a principal are specified
    in the Principal element of the Task Scheduler schema.

    If a task is registered using the at.exe command line tool, and this object is used to retrieve
    information about the task, then the LogonType property will return 0, the RunLevel property
    will return 0, and the UserId property will return Nothing.
    """

    @property
    def display_name(self) -> str:
        """For scripting, gets or sets the name of the principal.

        Remarks
        When reading or writing XML for a task, the display name for a principal is specified in
        the DisplayName element of the Task Scheduler schema.

        When setting this property value, the value can be text that is retrieved from a resource
        .dll file. A specialized string is used to reference the text from the resource file. The
        format of the string is $(@ [Dll], [ResourceID]) where [Dll] is the path to the .dll file
        that contains the resource and [ResourceID] is the identifier for the resource text. For
        example, setting this property value to $(@ %SystemRoot%\\System32\\ResourceName.dll, -101)
        will set the property to the value of the resource text with an identifier equal to -101 in
        the %SystemRoot%\\System32\\ResourceName.dll file.

        :returns: The name of the principal.
        """
        return self._obj.DisplayName

    @display_name.setter
    def display_name(self, value: str) -> None:
        self._obj.DisplayName = value

    @property
    def group_id(self) -> str:
        """For scripting, gets or sets the identifier of the user group that is required to run the
        tasks that are associated with the principal.

        Remarks

        Do not set this property if a user identifier is specified in the UserId property.

        When reading or writing XML for a task, the group identifier for a principal is specified
        in the GroupId element of the Task Scheduler schema.

        :returns: The identifier of the user group that is associated with this principal.
        """
        return self._obj.GroupId

    @group_id.setter
    def group_id(self, value: str) -> None:
        self._obj.GroupId = value

    @property
    def id(self) -> str:
        """For scripting, gets or sets the identifier of the principal.

        Remarks

        This identifier is also used when specifying the ActionCollection.Context property.

        When reading or writing XML for a task, the identifier of the principal is specified in the
        ID attribute of the Principal element of the Task Scheduler schema.

        :returns: The identifier of the principal.
        """
        return self._obj.Id

    @id.setter
    def id(self, value: str) -> None:
        self._obj.Id = value

    @property
    def logon_type(self) -> LogonType:
        """For scripting, gets or sets the security logon method that is required to run the tasks
        that are associated with the principal.

        Remarks

        This property is valid only when a user identifier is specified by the UserId property.

        When reading or writing XML for a task, the logon type is specified in the <LogonType>
        element of the Task Scheduler schema.

        For a task, that contains a message box action, the message box will be displayed if the
        task is activated and the task has an interactive logon type. To set the task logon type to
        interactive, specify 3 (TASK_LOGON_INTERACTIVE_TOKEN) or 4 (TASK_LOGON_GROUP) in the
        LogonType property of the task principal, or in the logonType parameter of
        TaskFolder.RegisterTask or TaskFolder.RegisterTaskDefinition.

        :returns: Set to one of the following TASK_LOGON TYPE enumeration constants.
        """
        return LogonType(self._obj.LogonType)

    @logon_type.setter
    def logon_type(self, value: LogonType) -> None:
        self._obj.LogonType = value.value

    @property
    def run_level(self) -> RunLevel:
        """For scripting, gets or sets the identifier that is used to specify the privilege level
        that is required to run the tasks that are associated with the principal.

        This property is read/write.

        Remarks

        If a task is registered using the Builtin/Administrator account or the Local System or
        Local Service accounts, then the RunLevel property will be ignored. The property value will
        also be ignored if User Account Control (UAC) is turned off.

        If a task is registered using the Administrators group for the security context of the task,
        then you must also set the RunLevel property to TASK_RUNLEVEL_HIGHEST if you want to run
        the task. For more information, see Security Contexts for Tasks.

        :returns: The identifier that is used to specify the privilege level that is required to
          run the tasks that are associated with the principal.
        """
        return RunLevel(self._obj.RunLevel)

    @run_level.setter
    def run_level(self, value: RunLevel) -> None:
        self._obj.RunLevel = value.value

    @property
    def user_id(self) -> str:
        """For scripting, gets or sets the user identifier that is required to run the tasks that
        are associated with the principal.

        Remarks

        Do not set this property if a group identifier is specified in the GroupId property.

        When reading or writing XML for a task, the user identifier for the principal is specified
        using the UserId element of the Task Scheduler schema.

        :returns: The user identifier that is required to run the task.
        """
        return self._obj.UserId

    @user_id.setter
    def user_id(self, value: str) -> None:
        self._obj.UserId = value


class TaskNamedValuePair(WrapperClass, Iterable, Sized):
    """Scripting object that is used to create a name-value pair in which the name is associated
    with the value.

    Remarks

    When reading or writing your own XML for a task, a name-value pair is specified using the
    ValueQueries element of the Task Scheduler schema.
    """

    def __len__(self) -> int:
        return 2

    def __getitem__(self, index: int) -> str:
        if index == 0:
            return self.name
        if index == 1:
            return self.value
        raise IndexError(f"Invalid index for TaskNamedValuePair: {index}")

    def __iter__(self) -> Iterator[str]:
        return iter((self.name, self.value))

    @property
    def name(self) -> str:
        """For scripting, gets or sets the name that is associated with a value in a name-value
        pair.

        :returns: The name that is associated with a value in a name-value pair.
        """
        return self._obj.Name

    @name.setter
    def name(self, value: str) -> None:
        self._obj.Name = value

    @property
    def value(self) -> str:
        """For scripting, gets or sets the value that is associated with a name in a name-value
        pair.

        :returns: The value that is associated with a name in a name-value pair.
        """
        return self._obj.Value

    @value.setter
    def value(self, value: str) -> None:
        self._obj.Value = value


class TaskNamedValueCollection(WrapperClass, Iterable, Sized):
    """Scripting object that contains a collection of TaskNamedValuePair object name-value pairs."""

    def __len__(self) -> int:
        return self.count

    def __getitem__(self, index: int) -> TaskNamedValuePair:
        return self.item(index)

    def __iter__(self) -> Iterator[TaskNamedValuePair]:
        return iter(map(TaskNamedValuePair, self._obj))

    @property
    def count(self) -> int:
        """For scripting, gets the number of name-value pairs in the collection.

        :returns: The number of name-value pairs in the collection.
        """
        return self._obj.Count

    def item(self, index: int) -> TaskNamedValuePair:
        """For scripting, gets the specified name-value pair from the collection.

        Remarks

        Collections are 1-based. In other words, the index for the first item in the collection is
        1.

        :returns: A TaskNamedValuePair object that represents the requested pair.
        """
        try:
            return TaskNamedValuePair(self._obj.Item(index))
        except com_error:
            raise IndexError(f"No TaskNamedValuePair at index: {index}") from None

    def clear(self) -> None:
        """For scripting, clears the entire collection of name-value pairs."""
        self._obj.Clear()

    def create(self, name: str, value: str) -> TaskNamedValuePair:
        """For scripting, creates a name-value pair in the collection.

        :param name: The name that is associated with a value in a name-value pair.
        :param value: The value that is associated with a name in a name-value pair.
        :returns: The name-value pair that is created in the collection.
        """
        return TaskNamedValuePair(self._obj.Create(name, value))

    def remove(self, index: int) -> None:
        """For scripting, removes a selected name-value pair from the collection.

        :param index: The index of the name-value pair to be removed.
        """
        self._obj.Remove(index)


# ---------- ACTION CLASSES ---------- #


class Action(WrapperClass, ABC):
    """Scripting object that provides the common properties that are inherited by all action
    objects. An action object is created by the ActionCollection.Create method.
    """

    @property
    def id(self) -> str:
        """For scripting, gets or sets the identifier of the action.

        For information on how actions and tasks work together, see Task Actions.

        :return: The user-defined identifier for the action. This identifier is used by the Task
          Scheduler for logging purposes.
        """
        return self._obj.Id

    @id.setter
    def id(self, value: str) -> None:
        self._obj.Id = value

    @functools.cached_property
    def type(self) -> ActionType:
        """For scripting, gets the type of the action.

        The action type is defined when the action is created and cannot be changed later. For
        information on creating an action, see ActionCollection.Create.

        For information on how actions and tasks work together, see Task Actions.

        :return: This property returns one of the following TASK_ACTION_TYPE enumeration constants.
        """
        return ActionType(self._obj.Type)


class ActionCollection(WrapperClass, Iterable, Sized):
    """Scripting object that contains the actions performed by the task.

    When reading or writing XML, the actions of a task are specified in the Actions element of the
    Task Scheduler schema.
    """

    @staticmethod
    def get_action_class(type: ActionType) -> Type[Action]:
        if type == ActionType.EXEC:
            return ExecAction
        elif type == ActionType.COM_HANDLER:
            return ComHandlerAction
        elif type == ActionType.SEND_EMAIL:
            return EmailAction
        elif type == ActionType.SHOW_MESSAGE:
            return ShowMessageAction
        else:
            raise RuntimeError(f"Invalid type: {type}")

    def __len__(self) -> int:
        return self.count

    def __getitem__(self, index: int) -> Action:
        return self.item(index)

    def __iter__(self) -> Iterator[Action]:
        def mapper(action_obj) -> Action:
            action_class: Type[Action] = self.get_action_class(ActionType(action_obj.Type))
            return action_class(action_obj)

        return iter(map(mapper, self._obj))

    @property
    def context(self) -> str:
        """For scripting, gets or sets the identifier of the principal for the task.

        The identifier of the principal for the task. The identifier that is specified here must
        match the identifier that is specified in the ID property of the IPrincipal interface that
        defines the principal.
        """
        return self._obj.Context

    @context.setter
    def context(self, value: str) -> None:
        self._obj.Context = value

    @property
    def count(self) -> int:
        """For scripting, gets the number of actions in the collection.

        The number of actions in the collection. The collection can contain up to 32 actions.
        """
        return self._obj.Count

    def item(self, index: int) -> Action:
        """For scripting, gets a specified action from the collection.

        Remarks

        Collections are 1-based. In other words, the index for the first item in the collection is
        1.

        :returns: An Action object that represents the requested action.
        """
        try:
            action_obj = self._obj.Item(index)
            action_class: Type[Action] = self.get_action_class(ActionType(action_obj.Type))
            return action_class(action_obj)
        except com_error:
            raise IndexError(f"No Action at index: {index}") from None

    def clear(self) -> None:
        """For scripting, clears all the actions from the collection."""
        self._obj.Clear()

    def create(self, type: ActionType) -> Action:
        """For scripting, creates and adds a new action to the collection.

        Remarks

        You cannot add more than 32 actions to the collection.

        :param type: This parameter is set to one of the following ActionType enumeration constants.
        :returns: An Action object that represents the new action.
        """
        action_class: Type[Action] = self.get_action_class(type)
        return action_class(self._obj.Create(type.value))

    def remove(self, index: int) -> None:
        """For scripting, removes the specified action from the collection.

        Remarks

        When removing items, note that the index for the first item in the collection is 1 and the
        index for the last item is the value of the ActionCollection.Count property.

        :param index: The index of the action to be removed.
        """
        self._obj.Remove(index)


class ExecAction(Action):
    """Scripting object that represents an action that executes a command-line operation.

    If environment variables are used in the Path, Arguments, or WorkingDirectory properties, then
    the values of the environment variables are cached and used when the Taskeng.exe (the task
    engine) is launched. Changes to the environment variables that occur after the task engine is
    launched will not be used by the task engine.

    This action performs a command-line operation. For example, the action could run a script or
    launch an executable.

    When reading or writing XML, an execution action is specified in the Exec element of the Task
    Scheduler schema.
    """

    @property
    def arguments(self) -> str:
        """For scripting, gets or sets the arguments associated with the command-line operation.

        When reading or writing XML, the command-line operation arguments are specified in the
        Arguments element of the Task Scheduler schema.

        :return: The arguments associated with the command-line operation
        """
        return self._obj.Arguments

    @arguments.setter
    def arguments(self, value: str) -> None:
        self._obj.Arguments = value

    @property
    def path(self) -> str:
        """For scripting, gets or sets the path to an executable file.

        This action performs a command-line operation. For example, the action could run a script
        or launch an executable.

        When reading or writing XML, the command-line operation path is specified in the Command
        element of the Task Scheduler schema.

        The path is checked to make sure it is valid when the task is registered, not when this
        property is set.

        :return: The path to an executable file.
        """
        return self._obj.Path

    @path.setter
    def path(self, value: str) -> None:
        self._obj.Path = value

    @property
    def working_directory(self) -> str:
        """For scripting, gets or sets the directory that contains either the executable file or
        the files that are used by the executable file.

        When reading or writing XML, the working directory of the application is specified in the
        WorkingDirectory element of the Task Scheduler schema.

        The path is checked to make sure it is valid when the task is registered, not when this
        property is set.

        :return: The directory that contains either the executable file or the files that are used
          by the executable file.
        """
        return self._obj.WorkingDirectory

    @working_directory.setter
    def working_directory(self, value: str) -> None:
        self._obj.WorkingDirectory = value


class ComHandlerAction(Action):
    """Scripting object that represents an action that fires a handler.

    COM handlers must implement the ITaskHandler interface for the Task Scheduler to start and stop
    the handler. In turn, the COM handler uses the methods of the TaskHandlerStatus object to
    communicate the status back to the Task Scheduler.

    When reading or writing XML, a COM handler action is specified in the ComHandler element of the
    Task Scheduler schema.
    """

    @property
    def class_id(self) -> str:
        """For scripting, gets or sets the identifier of the handler class.

        When reading or writing XML, the class of a COM handler is specified in the ClassId element
        of the Task Scheduler schema.

        :return: The identifier of the class that defines the handler to be fired.
        """
        return self._obj.ClassId

    @class_id.setter
    def class_id(self, value: str) -> None:
        self._obj.ClassId = value

    @property
    def data(self) -> str:
        """For scripting, gets or sets additional data that is associated with the handler.

        When reading or writing XML, the data of a COM handler is specified in the Data element of
        the Task Scheduler schema.

        :return: The arguments that are needed by the handler.
        """
        return self._obj.Data

    @data.setter
    def data(self, value: str) -> None:
        self._obj.Data = value


class EmailAction(Action):
    """[This object is no longer supported. Please use IExecAction with the powershell
    Send-MailMessage cmdlet as a workaround.]

    Scripting object that represents an action that sends an email message.

    Remarks

    The email action must have a valid value for the Server, From, and To or Cc properties for the
    task to register and run correctly.

    When reading or writing your own XML for a task, an email action is specified using the
    SendEmail element of the Task Scheduler schema.
    """

    @property
    def attachments(self) -> Optional[Sequence[str]]:
        """[This object is no longer supported. Please use IExecAction with the powershell
        Send-MailMessage cmdlet as a workaround.]

        For scripting, gets or sets an array of attachments that is sent with the email message.

        This property is read/write.

        Remarks

        A maximum of eight attachments can be in the array of attachments.

        :returns: An array of attachments that is sent with the email message.
        """
        return self._obj.Attachments

    @attachments.setter
    def attachments(self, value: Optional[Sequence[str]]) -> None:
        self._obj.Attachments = value

    @property
    def bcc(self) -> str:
        """[This object is no longer supported. Please use IExecAction with the powershell
        Send-MailMessage cmdlet as a workaround.]

        For scripting, gets or sets the email address or addresses that you want to Bcc in the email
        message.

        This property is read/write.

        :returns: The email address or addresses that you want to Bcc in the email message.
        """
        return self._obj.Bcc

    @bcc.setter
    def bcc(self, value: str) -> None:
        self._obj.Bcc = value

    @property
    def body(self) -> str:
        """[This object is no longer supported. Please use IExecAction with the powershell
        Send-MailMessage cmdlet as a workaround.]

        For scripting, gets or sets the body of the email that contains the email message.

        This property is read/write.

        Remarks
        When setting this property value, the value can be text that is retrieved from a resource
        .dll file. A specialized string is used to reference the text from the resource file. The
        format of the string is $(@ [Dll], [ResourceID]) where [Dll] is the path to the .dll file
        that contains the resource and [ResourceID] is the identifier for the resource text. For
        example, the setting this property value to $(@ %SystemRoot%\\System32\\ResourceName.dll,
        -101) will set the property to the value of the resource text with an identifier equal to
        -101 in the %SystemRoot%\\System32\\ResourceName.dll file.

        :returns: The body of the email that contains the email message.
        """
        return self._obj.Body

    @body.setter
    def body(self, value: str) -> None:
        self._obj.Body = value

    @property
    def cc(self) -> str:
        """[This object is no longer supported. Please use IExecAction with the powershell
        Send-MailMessage cmdlet as a workaround.]

        For scripting, gets or sets the email address or addresses that you want to Cc in the email
        message.

        This property is read/write.

        :returns: The email address or addresses that you want to Cc in the email message.
        """
        return self._obj.Cc

    @cc.setter
    def cc(self, value: str) -> None:
        self._obj.Cc = value

    @property
    def from_(self) -> str:
        """[This object is no longer supported. Please use IExecAction with the powershell
        Send-MailMessage cmdlet as a workaround.]

        For scripting, gets or sets the email address that you want to send the email from.

        This property is read/write.

        :returns: The email address that you want to send the email from.
        """
        return self._obj.From

    @from_.setter
    def from_(self, value: str) -> None:
        self._obj.From = value

    @functools.cached_property
    def header_fields(self) -> TaskNamedValueCollection:
        """[This object is no longer supported. Please use IExecAction with the powershell
        Send-MailMessage cmdlet as a workaround.]

        For scripting, gets or sets the header information in the email you want to send.

        This property is read/write.

        :returns: The header information in the email you want to send.
        """
        return TaskNamedValueCollection(self._obj.HeaderFields)

    @property
    def reply_to(self) -> str:
        """[This object is no longer supported. Please use IExecAction with the powershell
        Send-MailMessage cmdlet as a workaround.]

        For scripting, gets or sets the email address that you want to reply to.

        This property is read/write.

        :returns: The email address that you want to reply to.
        """
        return self._obj.ReplyTo

    @reply_to.setter
    def reply_to(self, value: str) -> None:
        self._obj.ReplyTo = value

    @property
    def server(self) -> str:
        """[This object is no longer supported. Please use IExecAction with the powershell
        Send-MailMessage cmdlet as a workaround.]

        For scripting, gets or sets the name of the SMTP server that you use to send email from.

        This property is read/write.

        Remarks

        Make sure the SMTP server that sends the email is set up correctly. E-mail is sent using
        NTLM authentication for Windows SMTP servers, which means that the security credentials
        used for running the task must also have privileges on the SMTP server to send email
        message. If the SMTP server is a non-Windows based server, then the email will be sent if
        the server allows anonymous access. For information about setting up the SMTP server, see
        SMTP Server Setup, and for information about managing SMTP server settings, see SMTP
        Administration.

        :returns: The name of the server that you use to send email from.
        """
        return self._obj.Server

    @server.setter
    def server(self, value: str) -> None:
        self._obj.Server = value

    @property
    def subject(self) -> str:
        """[This object is no longer supported. Please use IExecAction with the powershell
        Send-MailMessage cmdlet as a workaround.]

        For scripting, gets or sets the subject of the email message.

        This property is read/write.

        Remarks

        When setting this property value, the value can be text that is retrieved from a resource
        .dll file. A specialized string is used to reference the text from the resource file. The
        format of the string is $(@ [Dll], [ResourceID]) where [Dll] is the path to the .dll file
        that contains the resource and [ResourceID] is the identifier for the resource text. For
        example, the setting this property value to $(@ %SystemRoot%\\System32\\ResourceName.dll,
        -101) will set the property to the value of the resource text with an identifier equal to
        -101 in the %SystemRoot%\\System32\\ResourceName.dll file.

        :returns: The subject of the email message.
        """
        return self._obj.Subject

    @subject.setter
    def subject(self, value: str) -> None:
        self._obj.Subject = value

    @property
    def to(self) -> str:
        """[This object is no longer supported. Please use IExecAction with the powershell
        Send-MailMessage cmdlet as a workaround.]

        For scripting, gets or sets the email address or addresses that you want to send the email
        to.

        This property is read/write.

        :returns: The email address or addresses that you want to send the email to.
        """
        return self._obj.To

    @to.setter
    def to(self, value: str) -> None:
        self._obj.To = value


class ShowMessageAction(Action):
    """[This object is no longer supported. You can use IExecAction with the Windows scripting
    MsgBox function to show a message in the user session.]

    For scripting, represents an action that shows a message box when a task is activated.

    Remarks

    For a task, that contains a message box action, the message box will be displayed if the task
    is activated and the task has an interactive logon type. To set the task logon type to
    interactive, specify 3 (TASK_LOGON_INTERACTIVE_TOKEN) or 4 (TASK_LOGON_GROUP) in the LogonType
    property of the task principal, or in the logonType parameter of TaskFolder.RegisterTask or
    TaskFolder.RegisterTaskDefinition.

    When reading or writing your own XML for a task, a message box action is specified using the
    ShowMessage element of the Task Scheduler schema.
    """

    @property
    def message_body(self) -> str:
        """[This object is no longer supported. You can use IExecAction with the Windows scripting
        MsgBox function to show a message in the user session.]

        For scripting, gets or sets the message text that is displayed in the body of the message
        box.

        Remarks

        Parameterized strings can be used in the message text of the message box. For more
        information, see the Examples section in EventTrigger.ValueQueries.

        When setting this property value, the value can be text that is retrieved from a resource
        .dll file. A specialized string is used to reference the text from the resource file. The
        format of the string is $(@ [Dll], [ResourceID]) where [Dll] is the path to the .dll file
        that contains the resource and [ResourceID] is the identifier for the resource text. For
        example, the setting this property value to $(@ %SystemRoot%\\System32\\ResourceName.dll,
        -101) will set the property to the value of the resource text with an identifier equal to
        -101 in the %SystemRoot%\\System32\\ResourceName.dll file.

        :returns: The message text that is displayed in the body of the message box.
        """
        return self._obj.MessageBody

    @message_body.setter
    def message_body(self, value: str) -> None:
        self._obj.MessageBody = value

    @property
    def title(self) -> str:
        """[This object is no longer supported. You can use IExecAction with the Windows scripting
        MsgBox function to show a message in the user session.]

        For scripting, gets or sets the title of the message box.

        Remarks

        Parameterized strings can be used in the title text of the message box. For more
        information, see the Examples section in EventTrigger.ValueQueries.

        When setting this property value, the value can be text that is retrieved from a resource
        .dll file. A specialized string is used to reference the text from the resource file. The
        format of the string is $(@ [Dll], [ResourceID]) where [Dll] is the path to the .dll file
        that contains the resource and [ResourceID] is the identifier for the resource text. For
        example, the setting this property value to $(@ %SystemRoot%\\System32\\ResourceName.dll,
        -101) will set the property to the value of the resource text with an identifier equal to
        -101 in the %SystemRoot%\\System32\\ResourceName.dll file.

        :returns: The title of the message box.
        """
        return self._obj.Title

    @title.setter
    def title(self, value: str) -> None:
        self._obj.Title = value


# ---------- TRIGGER CLASSES ---------- #


class Trigger(WrapperClass, ABC):
    """Scripting object that provides the common properties that are inherited by all
    trigger objects.

    Remarks

    The Task Scheduler provides the following individual objects for the different triggers that a
    task can use:

    * BootTrigger
    * DailyTrigger
    * EventTrigger
    * IdleTrigger
    * LogonTrigger
    * MonthlyDOWTrigger
    * MonthlyTrigger
    * RegistrationTrigger
    * TimeTrigger
    * WeeklyTrigger
    * SessionStateChangeTrigger

    When reading or writing XML, the triggers of a task are specified in the Triggers
    element of the Task Scheduler schema.
    """

    @property
    def enabled(self) -> bool:
        """For scripting, gets or sets a Boolean value that indicates whether the trigger is
        enabled.

        When reading or writing XML for a task, the enabled property is specified using the Enabled
        element of the Task Scheduler schema.

        :return: True if the trigger is enabled; otherwise, false. The default is true.
        """
        return self._obj.Enabled

    @enabled.setter
    def enabled(self, value: bool) -> None:
        self._obj.Enabled = value

    @property
    def end_boundary(self) -> datetime:
        """For scripting, gets or sets the date and time when the trigger is deactivated. The
        trigger cannot start the task after it is deactivated.

        When reading or writing XML for a task, the enabled property is specified using the
        EndBoundary element of the Task Scheduler schema.

        :return: The date and time when the trigger is deactivated. The date and time must be in
          the following format: YYYY-MM-DDTHH:MM:SS(+-)HH:MM. For example the date October 11th,
          2005 at 1:21:17 in the Pacific time zone would be written as 2005-10-11T13:21:17-08:00.
          The (+-)HH:MM section of the format describes the time zone as a certain number of hours
          ahead or behind Coordinated Universal Time (Greenwich Mean Time).
        """
        return from_date_str(self._obj.EndBoundary)

    @end_boundary.setter
    def end_boundary(self, value: datetime) -> None:
        self._obj.EndBoundary = to_date_str(value)

    @property
    def execution_time_limit(self) -> relativedelta:
        """For scripting, gets or sets the maximum amount of time that the task launched by the
        trigger is allowed to run.

        When reading or writing XML for a task, the execution time limit is specified in the
        ExecutionTimeLimit element of the Task Scheduler schema.

        :return: The maximum amount of time that the task launched by the trigger is allowed to
          run. The format for this string is PnYnMnDTnHnMnS, where nY is the number of years, nM is
          the number of months, nD is the number of days, "T" is the date/time separator, nH is the
          number of hours, nM is the number of minutes, and nS is the number of seconds (for
          example, PT5M specifies 5 minutes and P1M4DT2H5M specifies one month, four days, two
          hours, and five minutes).
        """
        return from_duration_str(self._obj.ExecutionTimeLimit)

    @execution_time_limit.setter
    def execution_time_limit(self, value: relativedelta) -> None:
        self._obj.ExecutionTimeLimit = to_duration_str(value)

    @property
    def id(self) -> str:
        """For scripting, gets or sets the identifier for the trigger.

        When reading or writing XML for a task, the trigger identifier is specified in the ID
        attribute of the individual trigger elements (for example, the ID attribute of the
        BootTrigger element) of the Task Scheduler schema.

        :return: The identifier for the trigger. This identifier is used by the Task Scheduler for
          logging purposes.
        """
        return self._obj.Id

    @id.setter
    def id(self, value: str) -> None:
        self._obj.Id = value

    @functools.cached_property
    def repetition(self) -> RepetitionPattern:
        """For scripting, gets or sets a value that indicates how often the task is run and how
        long the repetition pattern is repeated after the task is started.

        When reading or writing your own XML for a task, the repetition pattern for a trigger is
        specified in the Repetition element of the Task Scheduler schema.

        :return: A RepetitionPattern object that defines how often the task is run and how long the
          repetition pattern is repeated after the task is started.
        """
        return RepetitionPattern(self._obj.Repetition)

    @property
    def start_boundary(self) -> Optional[datetime]:
        """For scripting, gets or sets the date and time when the trigger is activated.

        When reading or writing XML for a task, the trigger start boundary is specified
        in the StartBoundary element of the Task Scheduler schema.

        :return: The date and time when the trigger is activated. The date and time
          must be in the following format: YYYY-MM-DDTHH:MM:SS(+-)HH:MM. For example the
          date October 11th, 2005 at 1:21:17 in the Pacific time zone would be written
          as 2005-10-11T13:21:17-08:00. The (+-)HH:MM section of the format describes
          the time zone as a certain number of hours ahead or behind Coordinated
          Universal Time (Greenwich Mean Time).
        """
        return from_date_str(self._obj.StartBoundary)

    @start_boundary.setter
    def start_boundary(self, value: Optional[datetime]) -> None:
        self._obj.StartBoundary = to_date_str(value)

    @property
    def type(self) -> TriggerType:
        """For scripting, gets the type of the trigger. The trigger type is defined when
        the trigger is created and cannot be changed later. For information on creating
        a trigger, see TriggerCollection.Create.
        """
        return TriggerType(self._obj.Type)


class TriggerCollection(WrapperClass, Iterable, Sized):
    """Scripting object that is used to add to, remove from, and retrieve the triggers of a task.

    Remarks

    When reading or writing XML for a task, the triggers for the task are specified in the Triggers
    element of the Task Scheduler schema.

    For information about each trigger type see Trigger Types.
    """

    @staticmethod
    def get_trigger_class(type: TriggerType) -> Type[Trigger]:
        if type == TriggerType.EVENT:
            return EventTrigger
        elif type == TriggerType.TIME:
            return TimeTrigger
        elif type == TriggerType.DAILY:
            return DailyTrigger
        elif type == TriggerType.WEEKLY:
            return WeeklyTrigger
        elif type == TriggerType.MONTHLY:
            return MonthlyTrigger
        elif type == TriggerType.MONTHLY_DOW:
            return MonthlyDOWTrigger
        elif type == TriggerType.IDLE:
            return IdleTrigger
        elif type == TriggerType.REGISTRATION:
            return RegistrationTrigger
        elif type == TriggerType.BOOT:
            return BootTrigger
        elif type == TriggerType.LOGON:
            return LogonTrigger
        elif type == TriggerType.SESSION_STATE_CHANGE:
            return SessionStateChangeTrigger
        # elif type == TriggerType.CUSTOM_TRIGGER:
        #     return (self._obj.Create(type.value))
        else:
            raise RuntimeError(f"Invalid type: {type}")

    def __len__(self) -> int:
        return self.count

    def __getitem__(self, index: int) -> Trigger:
        return self.item(index)

    def __iter__(self) -> Iterator[Trigger]:
        def mapper(trigger_obj) -> Trigger:
            trigger_class: Type[Trigger] = self.get_trigger_class(TriggerType(trigger_obj.Type))
            return trigger_class(trigger_obj)

        return iter(map(mapper, self._obj))

    @property
    def count(self) -> int:
        """For scripting, gets the number of triggers in the collection.

        :returns: The number of triggers in the collection.
        """
        return self._obj.Count

    def item(self, index: int) -> Trigger:
        """For scripting, gets the specified trigger from the collection.

        Remarks

        Collections are 1-based. In other words, the index for the first item in the collection is
        1.

        :returns: A Trigger object that represents the requested trigger.
        """
        try:
            trigger_obj = self._obj.Item(index)
            trigger_class: Type[Trigger] = self.get_trigger_class(TriggerType(trigger_obj.Type))
            return trigger_class(trigger_obj)
        except com_error:
            raise IndexError(f"No Trigger at index: {index}") from None

    def clear(self) -> None:
        """For scripting, clears all triggers from the collection."""
        self._obj.Clear()

    def create(self, type: TriggerType) -> Trigger:
        """For scripting, creates a new trigger for the task.

        Remarks

        For information about each trigger type see Trigger Types.

        :param type: This parameter is set to one of the following TriggerType enumeration
          constants.
        :returns: A Trigger object that represents the new trigger.
        """
        trigger_class: Type[Trigger] = self.get_trigger_class(type)
        return trigger_class(self._obj.Create(type.value))

    def remove(self, index: int) -> None:
        """For scripting, removes the specified trigger from the collection of triggers used by the
        task.

        Remarks

        When removing items, note that the index for the first item in the collection is 1 and the
        index for the last item is the value of the TriggerCollection.Count property.

        :param index: The index of the trigger to be removed.
        """
        self._obj.Remove(index)


class EventTrigger(Trigger):
    """Scripting object that represents a trigger that starts a task when a system event occurs.

    A maximum of 500 tasks with event subscriptions can be created. An event subscription that
    queries for a variety of events can be used to trigger a task that uses the same action in
    response to the events being logged.

    When reading or writing your own XML for a task, an event trigger is specified using the
    EventTrigger element of the Task Scheduler schema.
    """

    @property
    def delay(self) -> relativedelta:
        """For scripting, gets or sets a value that indicates the amount of time between when the
        event occurs and when the task is started. The format for this string is PnYnMnDTnHnMnS,
        where nY is the number of years, nM is the number of months, nD is the number of days, "T"
        is the date/time separator, nH is the number of hours, nM is the number of minutes, and nS
        is the number of seconds (for example, PT5M specifies 5 minutes and P1M4DT2H5M specifies
        one month, four days, two hours, and five minutes).

        When reading or writing your own XML for a task, the event delay is specified using the
        Delay element of the Task Scheduler schema.

        :return: A value that indicates the amount of time between when the event occurs and when
          the task is started.
        """
        return from_duration_str(self._obj.Delay)

    @delay.setter
    def delay(self, value: relativedelta) -> None:
        self._obj.Delay = to_duration_str(value)

    @property
    def subscription(self) -> str:
        """For scripting, gets or sets a query string that identifies the event that fires the
        trigger.

        When reading or writing your own XML for a task, the event subscription is specified using
        the Subscription element of the Task Scheduler schema.

        For more information about writing a query string for certain events, see Event Selection
        and Subscribing to Events.

        :return: A query string that identifies the event that fires the trigger.
        """
        return self._obj.Subscription

    @subscription.setter
    def subscription(self, value: str) -> None:
        self._obj.Subscription = value

    _value_queries: TaskNamedValueCollection = None

    @property
    def value_queries(self) -> TaskNamedValueCollection:
        """For scripting, gets or sets a collection of named XPath queries. Each query in the
        collection is applied to the last matching event XML that is returned from the subscription
        query specified in the Subscription property.

        The name of the query can be used as a variable in the following action properties:

        * ShowMessageAction.MessageBody
        * ShowMessageAction.Title
        * ExecAction.Arguments
        * ExecAction.WorkingDirectory
        * EmailAction.Server
        * EmailAction.Subject
        * EmailAction.To
        * EmailAction.Cc
        * EmailAction.Bcc
        * EmailAction.ReplyTo
        * EmailAction.From
        * EmailAction.Body
        * ComHandlerAction.Data

        The following code example strings show two name value pairs that can be used in a
        name-value collection. The values returned by the XPath queries can replace variables in an
        action's property. The values are referenced by name, with $(user) and $(machine), in the
        action property. For example, if the $(user) and $(machine) variables are used in the
        ShowMessageAction.MessageBody string, then the value of the XPath queries will replace the
        variables in the string.

        syntax

        name: user
        value: Event/UserData/SubjectUserName

        name: machine
        value: Event/UserData/MachineName

        For more information about writing a query string for certain events, see Event Selection
        and Subscribing to Events.

        :return: A collection of name-value pairs. Each name-value pair in the collection defines a
          unique name for a property value of the event that triggers the event trigger. The
          property value of the event is defined as an XPath event query. For more information
          about XPath event queries, see Event Selection.
        """
        if self._value_queries is None:
            self._value_queries = TaskNamedValueCollection(self._obj.ValueQueries)
        return self._value_queries

    @value_queries.setter
    def value_queries(self, value: TaskNamedValueCollection) -> None:
        self._value_queries = value
        self._obj.ValueQueries = value._obj


class TimeTrigger(Trigger):
    """Scripting object that represents a trigger that starts a task at a specific date
    and time.

    The StartBoundary element is a required element for time and calendar triggers (TimeTrigger and
    CalendarTrigger).

    When reading or writing XML for a task, an idle trigger is specified using the TimeTrigger
    element of the Task Scheduler schema.
    """

    @property
    def random_delay(self) -> relativedelta:
        """For scripting, gets or sets a delay time that is randomly added to the start time of the
        trigger.

        :return: The delay time that is randomly added to the start time of the trigger. The format
          for this string is P<days>DT<hours>H<minutes>M<seconds>S (for example, P2DT5S is a 2-day,
          5-second delay).
        """
        return from_duration_str(self._obj.RandomDelay)

    @random_delay.setter
    def random_delay(self, value: relativedelta) -> None:
        self._obj.RandomDelay = to_duration_str(value)


class DailyTrigger(TimeTrigger):
    """Scripting object that represents a trigger that starts a task based on a daily schedule.

    Remarks

    The time of day that the task is started is set by the StartBoundary property.

    An interval of 1 produces a daily schedule. An interval of 2 produces an every other day
    schedule and so on.

    When reading or writing your own XML for a task, a daily trigger is specified using the
    ScheduleByDay element of the Task Scheduler schema.
    """

    @property
    def days_interval(self) -> int:
        """For scripting, gets or sets the interval between the days in the schedule.

        Remarks

        An interval of 1 produces a daily schedule. An interval of 2 produces an every-other day
        schedule.

        When reading or writing your own XML for a task, the interval for a daily schedule is
        specified using the DaysInterval element of the Task Scheduler schema.

        :return: The interval between the days in the schedule.
        """
        return self._obj.DaysInterval

    @days_interval.setter
    def days_interval(self, value: int) -> None:
        self._obj.DaysInterval = value


class WeeklyTrigger(TimeTrigger):
    """Scripting object that represents a trigger that starts a task based on a weekly schedule.
    For example, the task starts at 8:00 A.M. on a specific day of the week every week or every
    other week.

    Remarks

    The time of day that the task is started is set by the StartBoundary property.

    When reading or writing your own XML for a task, a weekly trigger is specified using the
    ScheduleByWeek element of the Task Scheduler schema.
    """

    @property
    def days_of_week(self) -> DaysOfWeek:
        """For scripting, gets or sets the days of the week on which the task runs.

        :returns: A bitwise mask that indicates the days of the week on which the task runs.
        """
        return DaysOfWeek(self._obj.DaysOfWeek)

    @days_of_week.setter
    def days_of_week(self, value: DaysOfWeek) -> None:
        self._obj.DaysOfWeek = value.value

    @property
    def weeks_interval(self) -> int:
        """For scripting, gets or sets the interval between the weeks in the schedule.

        Remarks

        An interval of 1 produces a weekly schedule. An interval of 2 produces an every-other week
        schedule.

        When reading or writing your own XML for a task, the interval for a weekly schedule is
        specified using the WeeksInterval element of the Task Scheduler schema.

        :returns: The interval between the weeks in the schedule.
        """
        return self._obj.WeeksInterval

    @weeks_interval.setter
    def weeks_interval(self, value: int) -> None:
        self._obj.WeeksInterval = value


class MonthlyTrigger(TimeTrigger):
    """Scripting object that represents a trigger that starts a task based on a monthly schedule.
    For example, the task starts on specific days of specific months.

    Remarks

    The time of day that the task is started is set by the StartBoundary property.

    When reading or writing your own XML for a task, a monthly trigger is specified using the
    ScheduleByMonth element of the Task Scheduler schema.
    """

    @property
    def days_of_month(self) -> DaysOfMonth:
        """For scripting, gets or sets the days of the month during which the task runs.

        :returns: A bitwise mask that indicates the days of the month during which the task runs.
        """
        return DaysOfMonth(self._obj.DaysOfMonth)

    @days_of_month.setter
    def days_of_month(self, value: DaysOfMonth) -> None:
        self._obj.DaysOfMonth = value.value

    @property
    def months_of_year(self) -> MonthsOfYear:
        """For scripting, gets or sets the months of the year during which the task runs.

        :returns: A bitwise mask that indicates the months of the year during which the task runs.
        """
        return MonthsOfYear(self._obj.MonthsOfYear)

    @months_of_year.setter
    def months_of_year(self, value: MonthsOfYear) -> None:
        self._obj.MonthsOfYear = value.value

    @property
    def run_on_last_day_of_month(self) -> bool:
        """For scripting, gets or sets a Boolean value that indicates that the task runs on the
        last week of the month.

        :returns: True indicates that the task runs on the last week of the month; otherwise, False.
        """
        return self._obj.RunOnLastDayOfMonth

    @run_on_last_day_of_month.setter
    def run_on_last_day_of_month(self, value: bool) -> None:
        self._obj.RunOnLastDayOfMonth = value


class MonthlyDOWTrigger(TimeTrigger):
    """Scripting object that represents a trigger that starts a task on a monthly day-of-week
    schedule. For example, the task starts on every first Thursday, May through October.

    Remarks

    The time of day that the task is started is set by the StartBoundary property.

    When reading or writing XML for a task, a monthly day-of-week trigger is specified using the
    ScheduleByMonthDayOfWeek element of the Task Scheduler schema.
    """

    @property
    def days_of_week(self) -> DaysOfWeek:
        """For scripting, gets or sets the days of the week during which the task runs.

        :returns: A bitwise mask that indicates the days of the week during which the task runs.
        """
        return DaysOfWeek(self._obj.DaysOfWeek)

    @days_of_week.setter
    def days_of_week(self, value: DaysOfWeek) -> None:
        self._obj.DaysOfWeek = value.value

    @property
    def months_of_year(self) -> MonthsOfYear:
        """For scripting, gets or sets the months of the year during which the task runs.

        :returns: A bitwise mask that indicates the months of the year during which the task runs.
        """
        return MonthsOfYear(self._obj.MonthsOfYear)

    @months_of_year.setter
    def months_of_year(self, value: MonthsOfYear) -> None:
        self._obj.MonthsOfYear = value.value

    @property
    def run_on_last_week_of_month(self) -> bool:
        """For scripting, gets or sets a Boolean value that indicates that the task runs on the
        last week of the month.

        :returns: True indicates that the task runs on the last week of the month; otherwise, False.
        """
        return self._obj.RunOnLastWeekOfMonth

    @run_on_last_week_of_month.setter
    def run_on_last_week_of_month(self, value: bool) -> None:
        self._obj.RunOnLastWeekOfMonth = value

    @property
    def weeks_of_month(self) -> WeeksOfMonth:
        return WeeksOfMonth(self._obj.WeeksOfMonth)

    @weeks_of_month.setter
    def weeks_of_month(self, value: WeeksOfMonth) -> None:
        self._obj.WeeksOfMonth = value.value


class IdleTrigger(Trigger):
    """Scripting object that represents a trigger that starts a task when an idle condition occurs.
    For information about idle conditions, see Task Idle Conditions.

    An idle trigger will only trigger a task action if the computer goes into an idle state after
    the start boundary of the trigger.

    When reading or writing XML for a task, an idle trigger is specified using the IdleTrigger
    element of the Task Scheduler schema.

    If a task is triggered by an idle trigger, then the IdleSettings.WaitTimeout property is
    ignored.

    If the initial instance of a task with an idle trigger is still running, then the task is only
    launched once with no repetitions, even if multiple repetition is defined in the Repetition
    property. This behavior does not occur if the task stops by itself.
    """


class RegistrationTrigger(Trigger):
    """Scripting object that represents a trigger that starts a task when the task is registered or
    updated.

    When creating your own XML for a task, a registration trigger is specified using the
    RegistrationTrigger element of the Task Scheduler schema.

    If a task with a delayed registration trigger is registered, and the computer that the task is
    registered on is shutdown or restarted during the delay, before the task runs, then the task
    will not run and the delay will be lost.
    """

    @property
    def delay(self) -> relativedelta:
        """For scripting, gets or sets the amount of time between when the task is registered and
        when the task is started. The format for this string is PnYnMnDTnHnMnS, where nY is the
        number of years, nM is the number of months, nD is the number of days, "T" is the date/time
        separator, nH is the number of hours, nM is the number of minutes, and nS is the number of
        seconds (for example, PT5M specifies 5 minutes and P1M4DT2H5M specifies one month, four
        days, two hours, and five minutes).

        When reading or writing XML for a task, the boot delay is specified using the Delay element
        of the Task Scheduler schema.

        If a task with a delayed registration trigger is registered, and the computer that the task
        is registered on is shutdown or restarted during the delay, before the task runs, then the
        task will not run and the delay will be lost.

        :return: The amount of time between when the system is registered and when the task is
          started.
        """
        return from_duration_str(self._obj.Delay)

    @delay.setter
    def delay(self, value: relativedelta) -> None:
        self._obj.Delay = to_duration_str(value)


class BootTrigger(Trigger):
    """Scripting object that represents a trigger that starts a task when the system is booted.

    Remarks

    The Task Scheduler service is started when the operating system is booted, and boot trigger
    tasks are set to start when the Task Scheduler service starts.

    Only a member of the Administrators group can create a task with a boot trigger.

    When creating your own XML for a task, a boot trigger is specified using the BootTrigger
    element of the Task Scheduler schema.
    """

    @property
    def delay(self) -> relativedelta:
        """For scripting, gets or sets a value that indicates the amount of time between when the
        system is booted and when the task is started.

        When reading or writing your own XML for a task, the boot delay is specified using the
        Delay element of the Task Scheduler schema.

        :return: A value that indicates the amount of time between when the system is booted and
          when the task is started. The format for this string is PnYnMnDTnHnMnS, where nY is the
          number of years, nM is the number of months, nD is the number of days, "T" is the
          date/time separator, nH is the number of hours, nM is the number of minutes, and nS is
          the number of seconds (for example, PT5M specifies 5 minutes and P1M4DT2H5M specifies one
          month, four days, two hours, and five minutes).
        """
        return from_duration_str(self._obj.Delay)

    @delay.setter
    def delay(self, value: relativedelta) -> None:
        self._obj.Delay = to_duration_str(value)


class LogonTrigger(Trigger):
    """Scripting object that represents a trigger that starts a task when a user logs on. When the
    Task Scheduler service starts, all logged-on users are enumerated and any tasks registered with
    logon triggers that match the logged on user are run.

    If you want a task to be triggered when any member of a group logs on to the computer rather
    than when a specific user logs on, then do not assign a value to the LogonTrigger.UserId
    property. Instead, create a logon trigger with an empty LogonTrigger.UserId property and assign
    a value to the principal for the task using the Principal.GroupId property.

    When reading or writing XML for a task, a logon trigger is specified using the LogonTrigger
    element of the Task Scheduler schema.
    """

    @property
    def delay(self) -> relativedelta:
        """For scripting, gets or sets a value that indicates the amount of time between when the
        user logs on and when the task is started.

        When reading or writing XML for a task, the logon trigger delay is specified using the
        Delay element of the Task Scheduler schema.

        :return: A value that indicates the amount of time between when the user logs on and when
          the task is started. The format for this string is PnYnMnDTnHnMnS, where nY is the number
          of years, nM is the number of months, nD is the number of days, "T" is the date/time
          separator, nH is the number of hours, nM is the number of minutes, and nS is the number
          of seconds (for example, PT5M specifies 5 minutes and P1M4DT2H5M specifies one month,
          four days, two hours, and five minutes).
        """
        return from_duration_str(self._obj.Delay)

    @delay.setter
    def delay(self, value: relativedelta) -> None:
        self._obj.Delay = to_duration_str(value)

    @property
    def user_id(self) -> str:
        """For scripting, gets or sets the identifier of the user.

        This property can be in one of the following formats:

        * Username or SID: The task is started when the user logs on to the computer.
        * NULL: The task is started when any user logs on to the computer.

        If you want a task to be triggered when any member of a group logs on to the computer
        rather than when a specific user logs on, then do not assign a value to the
        LogonTrigger.UserId property. Instead, create a logon trigger with an empty
        LogonTrigger.UserId property and assign a value to the principal for the task using the
        Principal.GroupId property.

        When reading or writing XML for a task, the logon user identifier is specified using the
        UserId element of the Task Scheduler schema.

        :return: The identifier of the user. For example, "MyDomain\\MyName" or for a local
        account, "Administrator".
        """
        return self._obj.UserId

    @user_id.setter
    def user_id(self, value: str) -> None:
        self._obj.UserId = value


class SessionStateChangeTrigger(Trigger):
    """Scripting object that triggers tasks for console connect or disconnect, remote connect or
    disconnect, or workstation lock or unlock notifications.

    When reading or writing your own XML for a task, a session state change trigger is specified
    using the SessionStateChangeTrigger element of the Task Scheduler schema.
    """

    @property
    def delay(self) -> relativedelta:
        """For scripting, gets or sets a value that indicates how long of a delay takes place
        before a task is started after a Terminal Server session state change is detected. The
        format for this string is PnYnMnDTnHnMnS, where nY is the number of years, nM is the number
        of months, nD is the number of days, "T" is the date/time separator, nH is the number of
        hours, nM is the number of minutes, and nS is the number of seconds (for example, PT5M
        specifies 5 minutes and P1M4DT2H5M specifies one month, four days, two hours, and five
        minutes).

        :return: The delay that takes place before a task is started after a Terminal Server
          session state change is detected.
        """
        return from_duration_str(self._obj.Delay)

    @delay.setter
    def delay(self, value: relativedelta) -> None:
        self._obj.Delay = to_duration_str(value)

    @property
    def state_change(self) -> SessionStateChangeType:
        """For scripting, gets or sets the kind of Terminal Server session change that would
        trigger a task launch.

        :return: The kind of Terminal Server session change that triggers a task to launch.
        """
        return SessionStateChangeType(self._obj.StateChange)

    @state_change.setter
    def state_change(self, value: SessionStateChangeType) -> None:
        self._obj.StateChange = value.value

    @property
    def user_id(self) -> str:
        """For scripting, gets or sets the user for the Terminal Server session. When a session
        state change is detected for this user, a task is started.

        :return: The user for the Terminal Server session.
        """
        return self._obj.UserId

    @user_id.setter
    def user_id(self, value: str) -> None:
        self._obj.UserId = value


# ---------- SETTINGS CLASSES ---------- #


class TaskSettings(WrapperClass):
    """A scripting object that provides the settings that the Task Scheduler service uses to perform
    the task.

    By default, a task will be stopped 72 hours after it starts to run. You can change this by
    changing the ExecutionTimeLimit setting.

    When reading or writing XML for a task, the task settings are defined in the Settings element
    of the Task Scheduler schema.
    """

    @property
    def allow_demand_start(self) -> bool:
        """For scripting, gets or sets a Boolean value that indicates that the task can be started
        by using either the Run command or the Context menu.

        When this property is set to True, the task can be started independent of when any triggers
        start the task.

        When reading or writing XML for a task, this setting is specified in the AllowStartOnDemand
        element of the Task Scheduler schema.

        :return: If True, the task can be run by using the Run command or the Context menu. If
          False, the task cannot be run using the Run command or the Context menu. The default is
          True.
        """
        return self._obj.AllowDemandStart

    @allow_demand_start.setter
    def allow_demand_start(self, value: bool) -> None:
        self._obj.AllowDemandStart = value

    @property
    def allow_hard_terminate(self) -> bool:
        """For scripting, gets or sets a Boolean value that indicates that the task may be
        terminated by the Task Scheduler service using TerminateProcess. The service will try to
        close the running task by sending the WM_CLOSE notification, and if the task does not
        respond, the task will be terminated only if this property is set to true.

        When reading or writing XML for a task, this setting is specified in the AllowHardTerminate
        element of the Task Scheduler schema.

        :return: If True, the task can be terminated by using TerminateProcess. If False, the task
          cannot be terminated by using TerminateProcess.
        """
        return self._obj.AllowHardTerminate

    @allow_hard_terminate.setter
    def allow_hard_terminate(self, value: bool) -> None:
        self._obj.AllowHardTerminate = value

    @property
    def compatibility(self) -> Compatibility:
        """For scripting, gets or sets a value that indicates which version of Task Scheduler a
        task is compatible with.

        Task compatibility, which is set through the Compatibility property, should only be set to
        TASK_COMPATIBILITY_V1 if a task needs to be accessed or modified from a Windows XP, Windows
        Server 2003, or Windows 2000 computer. Otherwise, it is recommended that Task Scheduler 2.0
        compatibility be used because the task will have more features.

        Tasks compatible with the AT command can only have one time trigger.

        Tasks compatible with Task Scheduler 1.0 can only have a time trigger, a logon trigger, or
        a boot trigger, and the task can only have an executable action.

        For more information about task compatibility, see What's New in Task Scheduler and Tasks.

        :return: A value that indicates which version of Task Scheduler a task is compatible with.
        """
        return Compatibility(self._obj.Compatibility)

    @compatibility.setter
    def compatibility(self, value: Compatibility) -> None:
        self._obj.Compatibility = value.value

    @property
    def delete_expired_task_after(self) -> Optional[datetime]:
        """For scripting, gets or sets the amount of time that the Task Scheduler will wait before
        deleting the task after it expires. If no value is specified for this property, then the
        Task Scheduler service will not delete the task.

        A task expires after the end boundary has been exceeded for all triggers associated with
        the task. The end boundary for a trigger is specified by the EndBoundary property inherited
        by all trigger objects.

        When reading or writing XML for a task, this setting is specified in the
        DeleteExpiredTaskAfter (settingsType) element of the Task Scheduler schema.

        :return: A string that gets or sets the amount of time that the Task Scheduler will wait
          before deleting the task after it expires. The format for this string is PnYnMnDTnHnMnS,
          where nY is the number of years, nM is the number of months, nD is the number of days,
          "T" is the date/time separator, nH is the number of hours, nM is the number of minutes,
          and nS is the number of seconds (for example, PT5M specifies 5 minutes and P1M4DT2H5M
          specifies one month, four days, two hours, and five minutes).
        """
        return from_date_str(self._obj.DeleteExpiredTaskAfter)

    @delete_expired_task_after.setter
    def delete_expired_task_after(self, value: Optional[datetime]) -> None:
        self._obj.DeleteExpiredTaskAfter = to_date_str(value)

    @property
    def disallow_start_if_on_batteries(self) -> bool:
        """For scripting, gets or sets a Boolean value that indicates that the task will not be
        started if the computer is running on batteries.

        When reading or writing XML for a task, this setting is specified in the
        DisallowStartIfOnBatteries element of the Task Scheduler schema.

        :return: A Boolean value that indicates that the task will not be started if the computer
        is running on batteries. If True, the task will not be started if the computer is running
        on batteries. If False, the task will be started if the computer is running on batteries.
        The default is True.
        """
        return self._obj.DisallowStartIfOnBatteries

    @disallow_start_if_on_batteries.setter
    def disallow_start_if_on_batteries(self, value: bool) -> None:
        self._obj.DisallowStartIfOnBatteries = value

    @property
    def enabled(self) -> bool:
        """For scripting, gets or sets a Boolean value that indicates that the task is enabled. The
        task can be performed only when this setting is True.

        When reading or writing XML for a task, this setting is specified in the Enabled
        (settingsType) element of the Task Scheduler schema.

        :return: If True, the task is enabled. If False, the task is not enabled.
        """
        return self._obj.Enabled

    @enabled.setter
    def enabled(self, value: bool) -> None:
        self._obj.Enabled = value

    @property
    def execution_time_limit(self) -> Optional[relativedelta]:
        """For scripting, gets or sets the amount of time that is allowed to complete the task. By
        default, a task will be stopped 72 hours after it starts to run. You can change this by
        changing this setting.

        When reading or writing XML for a task, this setting is specified in the ExecutionTimeLimit
        element of the Task Scheduler schema.

        :return: The amount of time that is allowed to complete the task. The format for this
          string is PnYnMnDTnHnMnS, where nY is the number of years, nM is the number of months, nD
          is the number of days, "T" is the date/time separator, nH is the number of hours, nM is
          the number of minutes, and nS is the number of seconds (for example, PT5M specifies 5
          minutes and P1M4DT2H5M specifies one month, four days, two hours, and five minutes). A
          value of PT0S will enable the task to run indefinitely. When this parameter is set to
          Nothing, the execution time limit is infinite.
        """

        return from_duration_str(self._obj.ExecutionTimeLimit)

    @execution_time_limit.setter
    def execution_time_limit(self, value: Optional[relativedelta]) -> None:
        self._obj.ExecutionTimeLimit = to_duration_str(value)

    @property
    def hidden(self) -> bool:
        """For scripting, gets or sets a Boolean value that indicates that the task will not be
        visible in the UI. However, administrators can override this setting through the use of a
        "master switch" that makes all tasks visible in the UI.

        When reading or writing XML for a task, this setting is specified in the Hidden
        (settingsType) element of the Task Scheduler schema.

        :return: If True, the value indicates that the task will not be visible in the UI. If
          False, the task will be visible in the UI. The default is False.
        """
        return self._obj.Hidden

    @hidden.setter
    def hidden(self, value: bool) -> None:
        self._obj.Hidden = value

    @functools.cached_property
    def idle_settings(self) -> IdleSettings:
        """For scripting, gets or sets the information that specifies how the Task Scheduler
        performs tasks when the computer is in an idle condition. For information about idle
        conditions, see Task Idle Conditions.

        When reading or writing XML for a task, this setting is specified in the IdleSettings
        element of the Task Scheduler schema.

        :return: An IdleSettings object that specifies how the Task Scheduler handles the task when
          the computer goes into an idle condition.
        """
        return IdleSettings(self._obj.IdleSettings)

    @property
    def multiple_instances(self) -> InstancesPolicy:
        """For scripting, gets or sets the policy that defines how the Task Scheduler deals with
        multiple instances of the task.

        When reading or writing XML for a task, this setting is specified in the
        MultipleInstancesPolicy element of the Task Scheduler schema.
        """
        return InstancesPolicy(self._obj.MultipleInstances)

    @multiple_instances.setter
    def multiple_instances(self, value: InstancesPolicy) -> None:
        self._obj.MultipleInstances = value.value

    @functools.cached_property
    def network_settings(self) -> NetworkSettings:
        """For scripting, gets or sets the network settings object that contains a network profile
        identifier and name. If the RunOnlyIfNetworkAvailable property of TaskSettings is True and
        a network prop-file is specified in the NetworkSettings property, then the task will run
        only if the specified network profile is available.
        """
        return NetworkSettings(self._obj.NetworkSettings)

    @property
    def priority(self) -> int:
        """For scripting, gets or sets the priority level of the task.

        Priority level 0 is the highest priority, and priority level 10 is the lowest priority. The
        default value is 7. Priority levels 7 and 8 are used for background tasks, and priority
        levels 4, 5, and 6 are used for interactive tasks.

        The task's action is started in a process with a priority that is based on a Priority Class
        value. A Priority Level value (thread priority) is used for COM handler, message box, and
        email task actions. For more information about the Priority Class and Priority Level values,
        see Scheduling Priorities. The following table lists the possible values for the priority
        parameter, and the corresponding Priority Class and Priority Level values.

         Task priority | Priority Class              | Priority Level
        ---------------|-----------------------------|-------------------------------
         0             | REALTIME_PRIORITY_CLASS     | THREAD_PRIORITY_TIME_CRITICAL
         1             | HIGH_PRIORITY_CLASS         | THREAD_PRIORITY_HIGHEST
         2             | ABOVE_NORMAL_PRIORITY_CLASS | THREAD_PRIORITY_ABOVE_NORMAL
         3             | ABOVE_NORMAL_PRIORITY_CLASS | THREAD_PRIORITY_ABOVE_NORMAL
         4             | NORMAL_PRIORITY_CLASS       | THREAD_PRIORITY_NORMAL
         5             | NORMAL_PRIORITY_CLASS       | THREAD_PRIORITY_NORMAL
         6             | NORMAL_PRIORITY_CLASS       | THREAD_PRIORITY_NORMAL
         7             | BELOW_NORMAL_PRIORITY_CLASS | THREAD_PRIORITY_BELOW_NORMAL
         8             | BELOW_NORMAL_PRIORITY_CLASS | THREAD_PRIORITY_BELOW_NORMAL
         9             | IDLE_PRIORITY_CLASS         | THREAD_PRIORITY_LOWEST
         10            | IDLE_PRIORITY_CLASS         | THREAD_PRIORITY_IDLE

        When reading or writing XML for a task, this setting is specified in the Priority
        (settingsType) element of the Task Scheduler schema.

        :return: The priority level (0-10) of the task. The default is 7.
        """
        return self._obj.Priority

    @priority.setter
    def priority(self, value: int) -> None:
        if value < 0 or value > 10:
            raise ValueError(f"Invalid Priority. [0 - 10]: {value}")
        self._obj.Priority = value

    @property
    def restart_count(self) -> int:
        """For scripting, gets or sets the number of times that the Task Scheduler will attempt to
        restart the task.

        The number of times that the Task Scheduler will attempt to restart the task.

        :return: The number of times that the Task Scheduler will attempt to restart the
          task.
        """
        return self._obj.RestartCount

    @restart_count.setter
    def restart_count(self, value: int) -> None:
        self._obj.RestartCount = value

    @property
    def restart_interval(self) -> Optional[relativedelta]:
        """For scripting, gets or sets a value that specifies how long the Task Scheduler will
        attempt to restart the task.

        A value that specifies how long the Task Scheduler will attempt to restart the task. If
        this property is set, the RestartCount property must also be set. The format for this
        string is P<days>DT<hours>H<minutes>M<seconds>S (for example, "PT5M" is 5 minutes, "PT1H"
        is 1 hour, and "PT20M" is 20 minutes). The maximum time allowed is 31 days, and the minimum
        time allowed is 1 minute.

        :return: A value that specifies how long the Task Scheduler will attempt to restart the
          task. If this property is set, the RestartCount property must also be set. The format for
          this string is P<days>DT<hours>H<minutes>M<seconds>S (for example, "PT5M" is 5 minutes,
          "PT1H" is 1 hour, and "PT20M" is 20 minutes). The maximum time allowed is 31 days, and
          the minimum time allowed is 1 minute.
        """
        return from_duration_str(self._obj.RestartInterval)

    @restart_interval.setter
    def restart_interval(self, value: Optional[relativedelta]) -> None:
        self._obj.RestartInterval = to_duration_str(value)

    @property
    def run_only_if_idle(self) -> bool:
        """For scripting, gets or sets a Boolean value that indicates that the Task Scheduler will
        run the task only if the computer is in an idle condition.

        :return: If True, the property indicates that the Task Scheduler will run the task only if
          the computer is in an idle condition. The default is False.
        """
        return self._obj.RunOnlyIfIdle

    @run_only_if_idle.setter
    def run_only_if_idle(self, value: bool) -> None:
        self._obj.RunOnlyIfIdle = value

    @property
    def run_only_if_network_available(self) -> bool:
        """For scripting, gets or sets a Boolean value that indicates that the Task Scheduler will
        run the task only when a network is available.

        When reading or writing XML for a task, this setting is specified in the
        RunOnlyIfNetworkAvailable element of the Task Scheduler schema.

        :return: If True, the property indicates that the Task Scheduler will run the task only
          when a network is available. The default is False.
        """
        return self._obj.RunOnlyIfNetworkAvailable

    @run_only_if_network_available.setter
    def run_only_if_network_available(self, value: bool) -> None:
        self._obj.RunOnlyIfNetworkAvailable = value

    @property
    def start_when_available(self) -> bool:
        """For scripting, gets or sets a Boolean value that indicates that the Task Scheduler can
        start the task at any time after its scheduled time has passed.

        This property applies only to time-based tasks with an end boundary or time-based tasks
        that are set to repeat infinitely.

        Tasks that are started after the scheduled time has passed (because of the
        StartWhenAvailable property being set to True) are queued in the Task Scheduler service's
        queue of tasks, and they are started after a delay. The default delay is 10 minutes.

        When reading or writing XML for a task, this setting is specified in the StartWhenAvailable
        element of the Task Scheduler schema.

        :return: If True, the property indicates that the Task Scheduler can start the task at any
          time after its scheduled time has passed. The default is False.
        """
        return self._obj.StartWhenAvailable

    @start_when_available.setter
    def start_when_available(self, value: bool) -> None:
        self._obj.StartWhenAvailable = value

    @property
    def stop_if_going_on_batteries(self) -> bool:
        """For scripting, gets or sets a Boolean value that indicates that the task will be stopped
        if the computer is going onto batteries.

        :return: A Boolean value that indicates that the task will be stopped if the computer is
          going onto batteries. If True, the property indicates that the task will be stopped if
          the computer is going onto batteries. If False, the property indicates that the task will
          not be stopped if the computer is going onto batteries. The default is True. See Remarks
          for more details.
        """
        return self._obj.StopIfGoingOnBatteries

    @stop_if_going_on_batteries.setter
    def stop_if_going_on_batteries(self, value: bool) -> None:
        self._obj.StopIfGoingOnBatteries = value

    @property
    def wake_to_run(self) -> bool:
        """For scripting, gets or sets a Boolean value that indicates that the Task Scheduler will
        wake the computer when it is time to run the task.

        When the Task Scheduler service wakes the computer to run a task, the screen may remain off
        even though the computer is no longer in the sleep or hibernate mode. The screen will turn
        on when Windows Vista detects that a user has returned to use the computer.

        When reading or writing XML for a task, this setting is specified in the WakeToRun element
        of the Task Scheduler schema.

        :return: For scripting, gets or sets a Boolean value that indicates that the Task Scheduler
          will wake the computer when it is time to run the task.
        """
        return self._obj.WakeToRun

    @wake_to_run.setter
    def wake_to_run(self, value: bool) -> None:
        self._obj.WakeToRun = value

    @property
    def xml_text(self) -> str:
        """For scripting, gets or sets an XML-formatted definition of the task settings.

        :return: An XML-formatted definition of the task settings.
        """
        return self._obj.XmlText

    @xml_text.setter
    def xml_text(self, value: str) -> None:
        self._obj.XmlText = value


class IdleSettings(WrapperClass):
    """A scripting object that specifies how the Task Scheduler performs tasks when the computer is
    in an idle condition. For information about idle conditions, see Task Idle Conditions.

    When reading or writing XML for a task, this setting is specified in the IdleSettings element
    of the Task Scheduler schema.

    If a task is triggered by an idle trigger, then the IdleSettings.WaitTimeout property is
    ignored.

    Note:
    IdleSettings.IdleDuration and IdleSettings.WaitTimeout are deprecated.
    """

    @property
    def restart_on_idle(self) -> bool:
        """For scripting, gets or sets a Boolean value that indicates whether the task is restarted
        when the computer cycles into an idle condition more than once.

        This property is only used if the IdleSettings.StopOnIdleEnd property is set to True.

        When reading or writing XML for a task, this setting is specified in the RestartOnIdle
        element of the Task Scheduler schema.

        :return: A Boolean value that indicates whether the task must be restarted when the
          computer cycles into an idle condition more than once. The default is False.
        """
        return self._obj.RestartOnIdle

    @restart_on_idle.setter
    def restart_on_idle(self, value: bool) -> None:
        self._obj.RestartOnIdle = value

    @property
    def stop_on_idle_end(self) -> bool:
        """For scripting, gets or sets a Boolean value that indicates that the Task Scheduler will
        terminate the task if the idle condition ends before the task is completed.

        :return: A Boolean value that indicates that the Task Scheduler will terminate the task if
          the idle condition ends before the task is completed.
        """
        return self._obj.StopOnIdleEnd

    @stop_on_idle_end.setter
    def stop_on_idle_end(self, value: bool) -> None:
        self._obj.StopOnIdleEnd = value


class NetworkSettings(WrapperClass):
    """For scripting, provides the settings that the Task Scheduler service uses to obtain a
    network profile.

    When reading or writing your own XML for a task, network settings are specified using the
    NetworkSettings element of the Task Scheduler schema.
    """

    @property
    def id(self) -> str:
        """For scripting, gets or sets a GUID value that identifies a network profile.

        :return: A GUID value that identifies a network profile.
        """
        return self._obj.Id

    @id.setter
    def id(self, value: str) -> None:
        self._obj.Id = value

    @property
    def name(self) -> str:
        """For scripting, gets or sets the name of a network profile. The name is used for display
        purposes.

        :return: The name of a network profile.
        """
        return self._obj.Name

    @name.setter
    def name(self, value: str) -> None:
        self._obj.Name = value


# ---------- ENUMS ---------- #


class ActionType(Enum):
    EXEC: ActionType = 0
    """This action performs a command-line operation. For example, the action can run a script,
    launch an executable, or, if the name of a document is provided, find its associated
    application and launch the application with the document.
    """
    COM_HANDLER: ActionType = 5
    """This action fires a handler. This action can only be used if the task Compatibility property
    is set to TASK_COMPATIBILITY_V2.
    """
    SEND_EMAIL: ActionType = 6
    """This action sends email message. This action can only be used if the task Compatibility
    property is set to TASK_COMPATIBILITY_V2.
    """
    SHOW_MESSAGE: ActionType = 7
    """This action shows a message box. This action can only be used if the task Compatibility
    property is set to TASK_COMPATIBILITY_V2.
    """


class Compatibility(Enum):
    AT: Compatibility = 0
    """The task is compatible with the AT command."""
    V1: Compatibility = 1
    """The task is compatible with Task Scheduler 1.0."""
    V2: Compatibility = 2
    """The task is compatible with Task Scheduler 2.0."""
    V2_1: Compatibility = 3
    V2_2: Compatibility = 4
    V2_3: Compatibility = 5
    V2_4: Compatibility = 6


class Creation(Flag):
    VALIDATE_ONLY: Creation = 0x1
    """The Task Scheduler service checks the syntax of the XML that describes the task but does not
    register the task. This constant cannot be combined with the TASK_CREATE, TASK_UPDATE, or
    TASK_CREATE_OR_UPDATE values.
    """
    CREATE: Creation = 0x2
    """The Task Scheduler service registers the task as a new task."""
    UPDATE: Creation = 0x4
    """The Task Scheduler service registers the task as an updated version of an existing task.
    When a task with a registration trigger is updated, the task will execute after the update
    occurs.
    """
    CREATE_OR_UPDATE: Creation = CREATE | UPDATE
    """The Task Scheduler service either registers the task as a new task or as an updated version
    if the task already exists. Equivalent to TASK_CREATE | TASK_UPDATE.
    """
    DISABLE: Creation = 0x8
    """The Task Scheduler service registers the disabled task. A disabled task cannot run until it
    is enabled. For more information, see Enabled Property of ITaskSettings and Enabled Property of
    IRegisteredTask.
    """
    DONT_ADD_PRINCIPAL_ACE: Creation = 0x10
    """The Task Scheduler service is prevented from adding the allow access-control entry (ACE) for
    the context principal. When the ITaskFolder::RegisterTaskDefinition or
    ITaskFolder::RegisterTask functions are called with this flag to update a task, the Task
    Scheduler service does not add the ACE for the new context principal and does not remove the
    ACE from the old context principal.
    """
    IGNORE_REGISTRATION_TRIGGERS: Creation = 0x20
    """The Task Scheduler service creates the task, but ignores the registration triggers in the
    task. By ignoring the registration triggers, the task will not execute when it is registered
    unless a time-based trigger causes it to execute on registration.
    """


class Flags(Flag):
    HIDDEN: Flags = 0x1
    """Enumerates all tasks, including tasks that are hidden."""


class InstancesPolicy(Enum):
    PARALLEL: InstancesPolicy = 0
    """Starts new instance while an existing instance is running."""
    QUEUE: InstancesPolicy = 1
    """Starts a new instance of the task after all other instances of the task are complete."""
    IGNORE_NEW: InstancesPolicy = 2
    """Does not start a new instance if an existing instance of the task is running."""
    STOP_EXISTING: InstancesPolicy = 3
    """Stops an existing instance of the task before it starts a new instance."""


class LogonType(Enum):
    NONE: LogonType = 0
    """The logon method is not specified. Used for non-NT credentials."""
    PASSWORD: LogonType = 1
    """ Use a password for logging on the user. The password must be supplied at registration time.
    """
    S4U: LogonType = 2
    """The service will log the user on using Service For User (S4U), and the task will run in a
    non-interactive desktop. When an S4U logon is used, no password is stored by the system and
    there is no access to either the network or to encrypted files.
    """
    INTERACTIVE_TOKEN: LogonType = 3
    """User must already be logged on. The task will be run only in an existing interactive session.
    """
    GROUP: LogonType = 4
    """Group activation. The groupId field specifies the group."""
    SERVICE_ACCOUNT: LogonType = 5
    """Indicates that a Local System, Local Service, or Network Service account is being used as a
    security context to run the task.
    """
    INTERACTIVE_TOKEN_OR_PASSWORD: LogonType = 6
    """Not in use; currently identical to TASK_LOGON_PASSWORD.
    
    Windows 10, version 1511, Windows 10, version 1507, Windows 8.1, Windows Server 2012 R2,
    Windows 8, Windows Server 2012, Windows Vista and Windows Server 2008:  First use the
    interactive token. If the user is not logged on (no interactive token is available), then the
    password is used. The password must be specified when a task is registered. This flag is not
    recommended for new tasks because it is less reliable than TASK_LOGON_PASSWORD.
    """


class ProcessTokenIDType(Enum):
    NONE: ProcessTokenIDType = 0
    """No changes will be made to the process token groups list."""
    UNRESTRICTED: ProcessTokenIDType = 1
    """A task SID that is derived from the task name will be added to the process token groups list,
    and the token default discretionary access control list (DACL) will be modified to allow only
    the task SID and local system full control and the account SID read control.
    """
    DEFAULT: ProcessTokenIDType = 2
    """A Task Scheduler will apply default settings to the task process."""


class RunFlags(Flag):
    NO_FLAGS: RunFlags = 0
    """The task is run with all flags ignored."""
    AS_SELF: RunFlags = 0x1
    """The task is run as the user who is calling the Run method."""
    IGNORE_CONSTRAINTS: RunFlags = 0x2
    """The task is run regardless of constraints such as "do not run on batteries" or
    "run only if idle".
    """
    USE_SESSION_ID: RunFlags = 0x4
    """The task is run using a terminal server session identifier."""
    USER_SID: RunFlags = 0x8
    """The task is run using a security identifier."""


class RunLevel(Enum):
    LUA: RunLevel = 0
    """Tasks will be run with the least privileges."""
    HIGHEST: RunLevel = 1
    """Tasks will be run with the highest privileges."""


class SessionStateChangeType(Enum):
    CONSOLE_CONNECT: SessionStateChangeType = 1
    """Terminal Server console connection state change. For example, when you connect to a user
    session on the local computer by switching users on the computer.
    """
    CONSOLE_DISCONNECT: SessionStateChangeType = 2
    """Terminal Server console disconnection state change. For example, when you disconnect to a
    user session on the local computer by switching users on the computer.
    """
    REMOTE_CONNECT: SessionStateChangeType = 3
    """Terminal Server remote connection state change. For example, when a user connects to a user
    session by using the Remote Desktop Connection program from a remote computer.
    """
    REMOTE_DISCONNECT: SessionStateChangeType = 4
    """Terminal Server remote disconnection state change. For example, when a user disconnects from
    a user session while using the Remote Desktop Connection program from a remote computer.
    """
    SESSION_LOCK: SessionStateChangeType = 7
    """Terminal Server session locked state change. For example, this state change causes the task
    to run when the computer is locked.
    """
    SESSION_UNLOCK: SessionStateChangeType = 8
    """Terminal Server session unlocked state change. For example, this state change causes the
    task to run when the computer is unlocked.
    """


class State(Enum):
    UNKNOWN: State = 0
    """The state of the task is unknown."""
    DISABLED: State = 1
    """The task is registered but is disabled and no instances of the task are queued or running.
    The task cannot be run until it is enabled.
    """
    QUEUED: State = 2
    """Instances of the task are queued."""
    READY: State = 3
    """The task is ready to be executed, but no instances are queued or running."""
    RUNNING: State = 4
    """One or more instances of the task is running."""


class TriggerType(Enum):
    EVENT: TriggerType = 0
    """Triggers the task when a specific event occurs. For more information about event triggers,
    see IEventTrigger.
    """
    TIME: TriggerType = 1
    """Triggers the task at a specific time of day. For more information about time triggers, see
    ITimeTrigger.
    """
    DAILY: TriggerType = 2
    """Triggers the task on a daily schedule. For example, the task starts at a specific time every
    day, every other day, or every third day. For more information about daily triggers, see
    IDailyTrigger.
    """
    WEEKLY: TriggerType = 3
    """Triggers the task on a weekly schedule. For example, the task starts at 8:00 AM on a
    specific day every week or other week. For more information about weekly triggers, see
    IWeeklyTrigger.
    """
    MONTHLY: TriggerType = 4
    """Triggers the task on a monthly schedule. For example, the task starts on specific days of
    specific months. For more information about monthly triggers, see IMonthlyTrigger.
    """
    MONTHLY_DOW: TriggerType = 5
    """Triggers the task on a monthly day-of-week schedule. For example, the task starts on a
    specific days of the week, weeks of the month, and months of the year. For more information
    about monthly day-of-week triggers, see IMonthlyDOWTrigger.
    """
    IDLE: TriggerType = 6
    """Triggers the task when the computer goes into an idle state. For more information about idle
    triggers, see IIdleTrigger.
    """
    REGISTRATION: TriggerType = 7
    """Triggers the task when the task is registered. For more information about registration
    triggers, see IRegistrationTrigger.
    """
    BOOT: TriggerType = 8
    """Triggers the task when the computer boots. For more information about boot triggers, see
    IBootTrigger.
    """
    LOGON: TriggerType = 9
    """Triggers the task when a specific user logs on. For more information about logon triggers,
    see ILogonTrigger.
    """
    SESSION_STATE_CHANGE: TriggerType = 11
    """Triggers the task when a specific user session state changes. For more information about
    session state change triggers, see ISessionStateChangeTrigger.
    """
    CUSTOM_TRIGGER: TriggerType = 12


class SecurityInformation(Flag):
    """The SECURITY_INFORMATION data type identifies the object-related security information being
    set or queried. This security information includes:

    * The owner of an object
    * The primary group of an object
    * The discretionary access control list (DACL) of an object
    * The system access control list (SACL) of an object

    Some SECURITY_INFORMATION members work only with the SetNamedSecurityInfo function. These
    members are not returned in the structure returned by other security functions such as
    GetNamedSecurityInfo or ConvertStringSecurityDescriptorToSecurityDescriptor.

    Each item of security information is designated by a bit flag. Each bit flag can be one of the
    following values. For more information, see the SetSecurityAccessMask and
    QuerySecurityAccessMask functions.
    """

    OWNER: SecurityInformation = 0x00000001
    """The owner of an object."""
    GROUP: SecurityInformation = 0x00000002
    """The primary group of an object."""
    DACL: SecurityInformation = 0x00000004
    """The discretionary access control list (DACL) of an object."""
    SACL: SecurityInformation = 0x00000008
    """The system access control list (SACL) of an object."""
    LABEL: SecurityInformation = 0x00000010


# ---------- PYTHON ENUMS ---------- #


class DaysOfMonth(Flag):
    FIRST: DaysOfMonth = 0x1
    SECOND: DaysOfMonth = 0x2
    THIRD: DaysOfMonth = 0x4
    FOURTH: DaysOfMonth = 0x8
    FIFTH: DaysOfMonth = 0x10
    SIXTH: DaysOfMonth = 0x20
    SEVENTH: DaysOfMonth = 0x40
    EIGHTH: DaysOfMonth = 0x80
    NINTH: DaysOfMonth = 0x100
    TENTH: DaysOfMonth = 0x200
    ELEVENTH: DaysOfMonth = 0x400
    TWELFTH: DaysOfMonth = 0x800
    THIRTEENTH: DaysOfMonth = 0x1000
    FOURTEENTH: DaysOfMonth = 0x2000
    FIFTEENTH: DaysOfMonth = 0x4000
    SIXTEENTH: DaysOfMonth = 0x8000
    SEVENTEENTH: DaysOfMonth = 0x10000
    EIGHTEENTH: DaysOfMonth = 0x20000
    NINETEENTH: DaysOfMonth = 0x40000
    TWENTIETH: DaysOfMonth = 0x80000
    TWENTY_FIRST: DaysOfMonth = 0x100000
    TWENTY_SECOND: DaysOfMonth = 0x200000
    TWENTY_THIRD: DaysOfMonth = 0x400000
    TWENTY_FOURTH: DaysOfMonth = 0x800000
    TWENTY_FIFTH: DaysOfMonth = 0x1000000
    TWENTY_SIXTH: DaysOfMonth = 0x2000000
    TWENTY_SEVENTH: DaysOfMonth = 0x4000000
    TWENTY_EIGHTH: DaysOfMonth = 0x8000000
    TWENTY_NINTH: DaysOfMonth = 0x10000000
    THIRTIETH: DaysOfMonth = 0x20000000
    THIRTY_FIRST: DaysOfMonth = 0x40000000
    LAST: DaysOfMonth = 0x80000000


class DaysOfWeek(Flag):
    SUNDAY: DaysOfWeek = 0x1
    MONDAY: DaysOfWeek = 0x2
    TUESDAY: DaysOfWeek = 0x4
    WEDNESDAY: DaysOfWeek = 0x8
    THURSDAY: DaysOfWeek = 0x10
    FRIDAY: DaysOfWeek = 0x20
    SATURDAY: DaysOfWeek = 0x40


class MonthsOfYear(Flag):
    JANUARY: MonthsOfYear = 0x1
    FEBRUARY: MonthsOfYear = 0x2
    MARCH: MonthsOfYear = 0x4
    APRIL: MonthsOfYear = 0x8
    MAY: MonthsOfYear = 0x10
    JUNE: MonthsOfYear = 0x20
    JULY: MonthsOfYear = 0x40
    AUGUST: MonthsOfYear = 0x80
    SEPTEMBER: MonthsOfYear = 0x100
    OCTOBER: MonthsOfYear = 0x200
    NOVEMBER: MonthsOfYear = 0x400
    DECEMBER: MonthsOfYear = 0x800


class WeeksOfMonth(Flag):
    FIRST: WeeksOfMonth = 0x1
    SECOND: WeeksOfMonth = 0x2
    THIRD: WeeksOfMonth = 0x4
    FOURTH: WeeksOfMonth = 0x8


# ---------- PYTHON EXCEPTIONS ---------- #


class TaskFolderNotFound(Exception):
    pass


class TaskFolderExists(Exception):
    pass


class TaskNotFound(Exception):
    pass


class TaskExists(Exception):
    pass


# ---------- USEFUL FUNCTIONS ---------- #


def xml_time(
    dt: datetime = None,
    /,
    *,
    year: int = None,
    month: int = None,
    day: int = None,
    hour: int = None,
    minute: int = None,
    second: int = None,
) -> str:
    """Get the time for the trigger start_boundary and end_boundary.

    :param dt: The datetime to convert
    :param year: The year part of the date
    :param month: The month part of the date
    :param day: The day part of the date
    :param hour: The hour part of the date
    :param minute: The minute part of the date
    :param second: The second part of the date
    :returns: The datetime in the correct format YYYY-MM-DDTHH:MM:SS
    """
    if dt is None:
        dt = datetime.now().replace(
            year=year,
            month=month,
            day=day,
            hour=hour,
            minute=minute,
            second=second,
        )
    return dt.isoformat()


def from_date_str(string: str) -> Optional[datetime]:
    if string == "":
        return None
    return datetime.fromisoformat(string)


def to_date_str(dt: Optional[datetime], default: str = "") -> str:
    if dt is None:
        return default
    return dt.isoformat()


# PnYnMnDTnHnMnS
duration_pattern: re.Pattern = re.compile(
    r"P(?:(\d+)Y)?(?:(\d+)M)?(?:(\d+)D)?T(?:(\d+)H)?(?:(\d+)M)?(?:(\d+)S)?",
    re.IGNORECASE,
)


def from_duration_str(string: str) -> Optional[relativedelta]:
    if string == "":
        return None

    match: re.Match = duration_pattern.match(string)
    if match is None:
        return None

    years: int = 0 if match.group(1) is None else int(match.group(1))
    months: int = 0 if match.group(2) is None else int(match.group(2))
    days: int = 0 if match.group(3) is None else int(match.group(3))
    hours: int = 0 if match.group(4) is None else int(match.group(4))
    minutes: int = 0 if match.group(5) is None else int(match.group(5))
    seconds: int = 0 if match.group(6) is None else int(match.group(6))

    if sum((years, months, days, hours, minutes, seconds)) == 0:
        return None

    return relativedelta(
        years=years,
        months=months,
        days=days,
        hours=hours,
        minutes=minutes,
        seconds=seconds,
    )


def to_duration_str(rd: Optional[relativedelta], default: str = "PT0S") -> str:
    if rd is None:
        return default

    parts: List[str] = ["P"]

    years: int = rd.years
    if years > 0:
        parts.append(f"{years}Y")

    months: int = rd.months
    if months > 0:
        parts.append(f"{months}M")

    days: int = rd.days
    if days > 0:
        parts.append(f"{days}D")

    parts.append("T")

    hours: int = rd.hours
    if hours > 0:
        parts.append(f"{hours}H")

    minutes: int = rd.minutes
    if minutes > 0:
        parts.append(f"{minutes}M")

    seconds: int = rd.seconds
    if seconds > 0:
        parts.append(f"{seconds}S")

    if len(parts) == 2:
        return default

    return "".join(parts)
