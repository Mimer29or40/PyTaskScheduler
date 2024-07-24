import unittest
from datetime import datetime
from datetime import timedelta
from typing import Final
from typing import Iterator
from typing import Optional
from typing import Sequence
from typing import Tuple
from typing import cast

from dateutil.relativedelta import relativedelta

from task_scheduler import Action
from task_scheduler import ActionCollection
from task_scheduler import ActionType
from task_scheduler import BootTrigger
from task_scheduler import ComHandlerAction
from task_scheduler import Compatibility
from task_scheduler import Creation
from task_scheduler import DailyTrigger
from task_scheduler import DaysOfMonth
from task_scheduler import DaysOfWeek
from task_scheduler import EmailAction
from task_scheduler import EventTrigger
from task_scheduler import ExecAction
from task_scheduler import IdleSettings
from task_scheduler import IdleTrigger
from task_scheduler import InstancesPolicy
from task_scheduler import LogonTrigger
from task_scheduler import LogonType
from task_scheduler import MonthlyDOWTrigger
from task_scheduler import MonthlyTrigger
from task_scheduler import MonthsOfYear
from task_scheduler import NetworkSettings
from task_scheduler import Principal
from task_scheduler import RegisteredTask
from task_scheduler import RegisteredTaskCollection
from task_scheduler import RegistrationInfo
from task_scheduler import RegistrationTrigger
from task_scheduler import RepetitionPattern
from task_scheduler import RunLevel
from task_scheduler import RunningTask
from task_scheduler import RunningTaskCollection
from task_scheduler import SecurityInformation
from task_scheduler import SessionStateChangeTrigger
from task_scheduler import SessionStateChangeType
from task_scheduler import ShowMessageAction
from task_scheduler import State
from task_scheduler import TaskDefinition
from task_scheduler import TaskFolder
from task_scheduler import TaskFolderCollection
from task_scheduler import TaskFolderExists
from task_scheduler import TaskFolderNotFound
from task_scheduler import TaskNamedValueCollection
from task_scheduler import TaskNamedValuePair
from task_scheduler import TaskNotFound
from task_scheduler import TaskService
from task_scheduler import TaskSettings
from task_scheduler import TimeTrigger
from task_scheduler import Trigger
from task_scheduler import TriggerCollection
from task_scheduler import TriggerType
from task_scheduler import WeeklyTrigger
from task_scheduler import WeeksOfMonth
from task_scheduler import from_duration_str
from task_scheduler import to_duration_str

SERVICE: Final[TaskService] = TaskService()
SERVICE.connect()

ROOT = SERVICE.get_folder("\\")


class TestTaskService(unittest.TestCase):
    def test_same_instance(self) -> None:
        new_service = TaskService()

        self.assertIs(SERVICE, new_service)

    def test_connected(self) -> None:
        connected: bool = SERVICE.connected
        self.assertIsInstance(connected, bool)

        connected_domain: str = SERVICE.connected_domain
        self.assertIsInstance(connected_domain, str)

        connected_user: str = SERVICE.connected_user
        self.assertIsInstance(connected_user, str)

        highest_version: int = SERVICE.highest_version
        self.assertIsInstance(highest_version, int)

        target_server: str = SERVICE.target_server
        self.assertIsInstance(target_server, str)

    def test_get_folder(self) -> None:
        folder: TaskFolder = SERVICE.get_folder("\\")
        self.assertIsInstance(folder, TaskFolder)

        self.assertRaises(TaskFolderNotFound, SERVICE.get_folder, "\\Folder\\Does\\Not\\Exist")

    def test_get_running_task(self) -> None:
        tasks: RunningTaskCollection = SERVICE.get_running_tasks()
        self.assertIsInstance(tasks, RunningTaskCollection)

        hidden_tasks: RunningTaskCollection = SERVICE.get_running_tasks(True)
        self.assertIsInstance(hidden_tasks, RunningTaskCollection)

    def test_new_task(self) -> None:
        new_task: TaskDefinition = SERVICE.new_task()
        self.assertIsInstance(new_task, TaskDefinition)


class TestTaskFolder(unittest.TestCase):
    def test_eq(self) -> None:
        folder: TaskFolder = ROOT.create_folder("TestEq")
        try:
            other: TaskFolder = SERVICE.get_folder(folder.path)
            self.assertEqual(folder, other)
        finally:
            ROOT.delete_folder(folder.path)

    def test_properties(self) -> None:
        folder: TaskFolder = SERVICE.get_folder("\\")

        name: str = folder.name
        self.assertIsInstance(name, str)

        path: str = folder.path
        self.assertIsInstance(path, str)

    def test_create_folder(self) -> None:
        folder: TaskFolder = ROOT.create_folder("TestCreate")
        try:
            self.assertIsInstance(folder, TaskFolder)
            self.assertEqual("TestCreate", folder.name)
            self.assertEqual("\\TestCreate", folder.path)
        finally:
            ROOT.delete_folder("TestCreate")

    def test_create_folder_nested(self) -> None:
        folder: TaskFolder = ROOT.create_folder("TestNested\\Folder")
        try:
            self.assertIsInstance(folder, TaskFolder)
            self.assertEqual("Folder", folder.name)
            self.assertEqual("\\TestNested\\Folder", folder.path)
        finally:
            ROOT.delete_folder("TestNested\\Folder")
            ROOT.delete_folder("TestNested")

    def test_create_folder_absolute(self) -> None:
        folder: TaskFolder = ROOT.create_folder("\\TestCreateAbsolute")
        try:
            self.assertIsInstance(folder, TaskFolder)
            self.assertEqual("TestCreateAbsolute", folder.name)
            self.assertEqual("\\TestCreateAbsolute", folder.path)
        finally:
            ROOT.delete_folder("\\TestCreateAbsolute")

    def test_create_folder_absolute_nested(self) -> None:
        folder: TaskFolder = ROOT.create_folder("\\TestAbsoluteNested\\Folder")
        try:
            self.assertIsInstance(folder, TaskFolder)
            self.assertEqual("Folder", folder.name)
            self.assertEqual("\\TestAbsoluteNested\\Folder", folder.path)
        finally:
            ROOT.delete_folder("\\TestAbsoluteNested\\Folder")
            ROOT.delete_folder("\\TestAbsoluteNested")

    def test_create_folder_exists(self) -> None:
        ROOT.create_folder("TestCreateExists")
        try:
            self.assertRaises(TaskFolderExists, ROOT.create_folder, "TestCreateExists")
        finally:
            ROOT.delete_folder("TestCreateExists")

    def test_create_folder_security_descriptor(self) -> None:  # TODO
        pass

    def test_delete_folder(self) -> None:
        folder: TaskFolder = ROOT.create_folder("TestDelete")
        try:
            self.assertIsNotNone(SERVICE.get_folder(folder.path))
            ROOT.delete_folder("TestDelete")
            self.assertRaises(TaskFolderNotFound, SERVICE.get_folder, folder.path)
        finally:
            try:
                ROOT.delete_folder("TestDelete")
            except TaskFolderNotFound:
                pass

    def test_delete_folder_nested(self) -> None:
        folder: TaskFolder = ROOT.create_folder("TestDeleteNested\\Folder")
        try:
            self.assertIsNotNone(SERVICE.get_folder(folder.path))
            ROOT.delete_folder("TestDeleteNested\\Folder")
            self.assertRaises(TaskFolderNotFound, SERVICE.get_folder, folder.path)
        finally:
            try:
                ROOT.delete_folder("TestDeleteNested\\Folder")
            except TaskFolderNotFound:
                pass
            try:
                ROOT.delete_folder("TestDeleteNested")
            except TaskFolderNotFound:
                pass

    def test_delete_folder_absolute(self) -> None:
        folder: TaskFolder = ROOT.create_folder("\\TestDeleteAbsolute")
        try:
            self.assertIsNotNone(SERVICE.get_folder(folder.path))
            ROOT.delete_folder("\\TestDeleteAbsolute")
            self.assertRaises(TaskFolderNotFound, SERVICE.get_folder, folder.path)
        finally:
            try:
                ROOT.delete_folder("\\TestDeleteAbsolute")
            except TaskFolderNotFound:
                pass

    def test_delete_folder_absolute_nested(self) -> None:
        folder: TaskFolder = ROOT.create_folder("\\TestDeleteAbsoluteNested\\Folder")
        try:
            self.assertIsNotNone(SERVICE.get_folder(folder.path))
            ROOT.delete_folder("\\TestDeleteAbsoluteNested\\Folder")
            self.assertRaises(TaskFolderNotFound, SERVICE.get_folder, folder.path)
        finally:
            try:
                ROOT.delete_folder("\\TestDeleteAbsoluteNested\\Folder")
            except TaskFolderNotFound:
                pass
            try:
                ROOT.delete_folder("\\TestDeleteAbsoluteNested")
            except TaskFolderNotFound:
                pass

    def test_delete_folder_not_found(self) -> None:
        self.assertRaises(TaskFolderNotFound, ROOT.delete_folder, "\\Folder\\Does\\Not\\Exist")

    def test_delete_task(self) -> None:  # TODO
        folder: TaskFolder = ROOT.create_folder("TestDeleteTask")

        # task_def: TaskDefinition = SERVICE.new_task()
        # task_def.registration_info.description = "Test Task"
        # task_def.settings.enabled = True
        # task_def.settings.stop_if_going_on_batteries = False
        #
        # start_time = datetime.now() + timedelta(minutes=5)
        # trigger: TimeTrigger = cast(TimeTrigger, task_def.triggers.create(TriggerType.TIME))
        # trigger.start_boundary = start_time.isoformat()
        #
        # # Create action
        # action: ExecAction = cast(ExecAction, task_def.actions.create(ActionType.EXEC))
        # action.id = "DO NOTHING"
        # action.path = "cmd.exe"
        # action.arguments = '/c "exit"'

        try:
            # task: RegisteredTask = folder.register_task_definition(
            #     "Test Task",  # Task name
            #     task_def,
            #     Creation.CREATE_OR_UPDATE,
            #     "",  # No user
            #     "",  # No password
            #     LogonType.NONE
            # )
            #
            # found: RegisteredTask = folder.get_task("Test Task")
            # folder.delete_task("Test Task")
            # self.assertRaises(TaskNotFound, folder.delete_task, "Test Task")
            pass
        finally:
            ROOT.delete_folder("TestDeleteTask")

    def test_get_folder(self) -> None:
        test_folder: TaskFolder = ROOT.create_folder("TestGet")
        folder: TaskFolder = test_folder.create_folder("Folder")
        try:
            found: TaskFolder = test_folder.get_folder("Folder")
            self.assertIsInstance(folder, TaskFolder)
            self.assertEqual(folder, found)

        finally:
            ROOT.delete_folder("TestGet\\Folder")
            ROOT.delete_folder("TestGet")

    def test_get_folder_not_found(self) -> None:
        test_folder: TaskFolder = ROOT.create_folder("TestGetNotFound")
        try:
            self.assertRaises(TaskFolderNotFound, test_folder.get_folder, "Folder")
        finally:
            ROOT.delete_folder("TestGetNotFound")

    def test_get_folders(self) -> None:
        try:
            folder: TaskFolder = ROOT.create_folder("TestGetMany")
            for i in range(5):
                folder.create_folder(f"Folder{i}")

            found: TaskFolderCollection = folder.get_folders()
            self.assertIsInstance(found, TaskFolderCollection)
        finally:
            for i in range(5):
                try:
                    ROOT.delete_folder(f"TestGetMany\\Folder{i}")
                except TaskFolderNotFound:
                    pass
            try:
                ROOT.delete_folder("TestGetMany")
            except TaskFolderNotFound:
                pass

    def test_get_folders_none(self) -> None:
        try:
            folder: TaskFolder = ROOT.create_folder("TestGetManyNone")
            found: TaskFolderCollection = folder.get_folders()
            self.assertIsInstance(found, TaskFolderCollection)
        finally:
            try:
                ROOT.delete_folder("TestGetManyNone")
            except TaskFolderNotFound:
                pass

    def test_get_security_descriptor(self) -> None:  # TODO
        pass

    def test_get_task(self) -> None:  # TODO
        pass

    def test_get_tasks(self) -> None:  # TODO
        pass

    def test_register_task(self) -> None:  # TODO
        pass

    def test_register_task_definition(self) -> None:  # TODO
        pass

    def test_set_security_description(self) -> None:  # TODO
        pass


class TestTaskFolderCollection(unittest.TestCase):
    folder: TaskFolder

    @classmethod
    def setUpClass(cls):
        try:
            cls.folder = ROOT.create_folder("TestFolderCollection")
        except TaskFolderExists:
            cls.folder = ROOT.get_folder("TestFolderCollection")

        for i in range(5):
            try:
                cls.folder.create_folder(f"Folder{i}")
            except TaskFolderExists:
                pass

    @classmethod
    def tearDownClass(cls):
        for i in range(5):
            try:
                cls.folder.delete_folder(f"Folder{i}")
            except TaskFolderNotFound:
                pass
        try:
            ROOT.delete_folder("TestFolderCollection")
        except TaskFolderNotFound:
            pass

    def test_dunder_len(self) -> None:
        collection: TaskFolderCollection = self.folder.get_folders()

        self.assertEqual(5, len(collection))

    def test_dunder_getitem(self) -> None:
        collection: TaskFolderCollection = self.folder.get_folders()

        folder: TaskFolder = collection[1]
        self.assertIsInstance(folder, TaskFolder)
        self.assertEqual("Folder0", folder.name)

    def test_dunder_getitem_zero(self) -> None:
        collection: TaskFolderCollection = self.folder.get_folders()

        self.assertRaises(IndexError, collection.__getitem__, 0)

    def test_dunder_getitem_out_of_range(self) -> None:
        collection: TaskFolderCollection = self.folder.get_folders()

        self.assertRaises(IndexError, collection.__getitem__, 6)

    def test_dunder_iter(self) -> None:
        collection: TaskFolderCollection = self.folder.get_folders()

        iterator: Iterator[TaskFolder] = iter(collection)
        self.assertIsInstance(iterator, Iterator)

        for member in collection:
            self.assertIsInstance(member, TaskFolder)

    def test_properties(self) -> None:
        collection: TaskFolderCollection = self.folder.get_folders()

        self.assertEqual(5, collection.count)

    def test_item(self) -> None:
        collection: TaskFolderCollection = self.folder.get_folders()

        folder: TaskFolder = collection.item(1)
        self.assertIsInstance(folder, TaskFolder)
        self.assertEqual("Folder0", folder.name)

    def test_item_zero(self) -> None:
        collection: TaskFolderCollection = self.folder.get_folders()

        self.assertRaises(IndexError, collection.item, 0)

    def test_item_out_of_range(self) -> None:
        collection: TaskFolderCollection = self.folder.get_folders()

        self.assertRaises(IndexError, collection.item, 6)


class TestTaskDefinition(unittest.TestCase):
    def test_actions(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()

        value: ActionCollection = task_def.actions
        self.assertIsInstance(value, ActionCollection)
        self.assertIs(value, task_def.actions)

    def test_data(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()

        value: str = task_def.data
        self.assertIsInstance(value, str)

    def test_principal(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()

        value: Principal = task_def.principal
        self.assertIsInstance(value, Principal)
        self.assertIs(value, task_def.principal)

    def test_registration_info(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()

        value: RegistrationInfo = task_def.registration_info
        self.assertIsInstance(value, RegistrationInfo)
        self.assertIs(value, task_def.registration_info)

    def test_settings(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()

        value: TaskSettings = task_def.settings
        self.assertIsInstance(value, TaskSettings)
        self.assertIs(value, task_def.settings)

    def test_triggers(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()

        value: TriggerCollection = task_def.triggers
        self.assertIsInstance(value, TriggerCollection)
        self.assertIs(value, task_def.triggers)

    def test_xml_text(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()

        value: str = task_def.xml_text
        self.assertIsInstance(value, str)


# class TestRunningTask(unittest.TestCase):  # TODO
# class TestRunningTaskCollection(unittest.TestCase):  # TODO


class TestRegisteredTask(unittest.TestCase):
    task: RegisteredTask

    @classmethod
    def setUpClass(cls):
        task_def: TaskDefinition = SERVICE.new_task()

        task_def.registration_info.description = "TestRegisteredTask"
        task_def.settings.enabled = True
        task_def.settings.stop_if_going_on_batteries = False

        trigger: TimeTrigger = cast(TimeTrigger, task_def.triggers.create(TriggerType.TIME))
        trigger.start_boundary = datetime.now() + timedelta(hours=1)

        action: ExecAction = cast(ExecAction, task_def.actions.create(ActionType.EXEC))
        action.id = "DO NOTHING"
        action.path = "cmd.exe"
        action.arguments = '/c "exit"'
        # action.arguments = '/c "timeout /t 30 > nul"'

        cls.task = ROOT.register_task_definition(
            "TestRegisteredTask",
            task_def,
            Creation.CREATE_OR_UPDATE,
            "",
            "",
            LogonType.NONE,
        )

    @classmethod
    def tearDownClass(cls):
        try:
            ROOT.delete_task("TestRegisteredTask")
        except TaskNotFound:
            pass

    def test_definition(self) -> None:
        value: TaskDefinition = self.task.definition
        self.assertIsInstance(value, TaskDefinition)
        self.assertIs(value, self.task.definition)

    def test_enabled(self) -> None:
        expected: bool = True
        self.task.enabled = expected

        value: bool = self.task.enabled
        self.assertIsInstance(value, bool)
        self.assertEqual(expected, value)

    def test_last_run_time(self) -> None:
        value: datetime = self.task.last_run_time
        self.assertIsInstance(value, datetime)

    def test_last_task_result(self) -> None:
        value: int = self.task.last_task_result
        self.assertIsInstance(value, int)

    def test_name(self) -> None:
        value: str = self.task.name
        self.assertIsInstance(value, str)

    def test_next_run_time(self) -> None:
        value: datetime = self.task.next_run_time
        self.assertIsInstance(value, datetime)

    def test_number_of_missed_runs(self) -> None:
        value: int = self.task.number_of_missed_runs
        self.assertIsInstance(value, int)

    def test_path(self) -> None:
        value: str = self.task.path
        self.assertIsInstance(value, str)

    def test_state(self) -> None:
        value: State = self.task.state
        self.assertIsInstance(value, State)

    def test_xml(self) -> None:
        value: str = self.task.xml
        self.assertIsInstance(value, str)

    def test_get_instances(self) -> None:
        value: RunningTaskCollection = self.task.get_instances()
        self.assertIsInstance(value, RunningTaskCollection)

    # def test_get_run_times(self) -> None:  # TODO
    #     start: datetime = datetime.now()
    #     end: datetime = start + relativedelta(years=1)
    #
    #     value: RunningTaskCollection = self.task.get_run_times(start, end)
    #     self.assertIsInstance(value, RunningTaskCollection)

    def test_get_security_descriptor(self) -> None:
        value: str = self.task.get_security_descriptor(
            SecurityInformation.OWNER
            | SecurityInformation.GROUP
            | SecurityInformation.DACL
            | SecurityInformation.SACL
            | SecurityInformation.LABEL
        )
        self.assertIsInstance(value, str)

    def test_run(self) -> None:
        value: RunningTask = self.task.run(None)
        self.assertIsInstance(value, RunningTask)

    # def test_run_ex(self) -> None:  # TODO
    #     value: Optional[RunningTask] = None
    #     try:
    #         value: RunningTask = self.task.run_ex(None, RunFlags.NO_FLAGS, 0)
    #         self.assertIsInstance(value, RunningTask)
    #     finally:
    #         if value is not None:
    #             value.stop()

    # def test_set_security_descriptor(self) -> None:  # TODO
    #     security_descriptor: str = self.task.get_security_descriptor(
    #         SecurityInformation.OWNER
    #         | SecurityInformation.GROUP
    #         | SecurityInformation.DACL
    #         | SecurityInformation.SACL
    #         | SecurityInformation.LABEL
    #     )
    #     flags: Creation = Creation.CREATE_OR_UPDATE
    #     self.task.set_security_descriptor(security_descriptor, flags)

    # def test_stop(self) -> None:  # TODO
    #     value: RunningTask = self.task.run(None)
    #     while value.state != State.READY:
    #         value.stop()


class TestRegisteredTaskCollection(unittest.TestCase):
    folder: TaskFolder

    @classmethod
    def setUpClass(cls):
        try:
            cls.folder = ROOT.create_folder("TestRegisteredTaskCollection")
        except TaskFolderExists:
            cls.folder = ROOT.get_folder("TestRegisteredTaskCollection")

        for i in range(5):
            task_def: TaskDefinition = SERVICE.new_task()

            task_def.registration_info.description = f"Test Task {i}"
            task_def.settings.enabled = True
            task_def.settings.stop_if_going_on_batteries = False

            trigger: TimeTrigger = cast(TimeTrigger, task_def.triggers.create(TriggerType.TIME))
            trigger.start_boundary = datetime.now() + timedelta(minutes=5)

            action: ExecAction = cast(ExecAction, task_def.actions.create(ActionType.EXEC))
            action.id = "DO NOTHING"
            action.path = "cmd.exe"
            action.arguments = '/c "exit"'

            cls.folder.register_task_definition(
                f"Test Task {i}",
                task_def,
                Creation.CREATE_OR_UPDATE,
                "",
                "",
                LogonType.NONE,
            )

    @classmethod
    def tearDownClass(cls):
        for i in range(5):
            try:
                cls.folder.delete_task(f"Test Task {i}")
            except TaskNotFound:
                pass
        try:
            ROOT.delete_folder("TestRegisteredTaskCollection")
        except TaskFolderNotFound:
            pass

    def test_dunder_len(self) -> None:
        collection: RegisteredTaskCollection = self.folder.get_tasks()

        self.assertEqual(5, len(collection))

    def test_dunder_getitem(self) -> None:
        collection: RegisteredTaskCollection = self.folder.get_tasks()

        task: RegisteredTask = collection[1]
        self.assertIsInstance(task, RegisteredTask)
        self.assertEqual("Test Task 0", task.name)

    def test_dunder_getitem_zero(self) -> None:
        collection: RegisteredTaskCollection = self.folder.get_tasks()

        self.assertRaises(IndexError, collection.__getitem__, 0)

    def test_dunder_getitem_out_of_range(self) -> None:
        collection: RegisteredTaskCollection = self.folder.get_tasks()

        self.assertRaises(IndexError, collection.__getitem__, 6)

    def test_dunder_iter(self) -> None:
        collection: RegisteredTaskCollection = self.folder.get_tasks()

        iterator: Iterator[RegisteredTask] = iter(collection)
        self.assertIsInstance(iterator, Iterator)

        for member in collection:
            self.assertIsInstance(member, RegisteredTask)

    def test_count(self) -> None:
        collection: RegisteredTaskCollection = self.folder.get_tasks()

        self.assertEqual(5, collection.count)

    def test_item(self) -> None:
        collection: RegisteredTaskCollection = self.folder.get_tasks()

        task: RegisteredTask = collection.item(1)
        self.assertIsInstance(task, RegisteredTask)
        self.assertEqual("Test Task 0", task.name)

    def test_item_zero(self) -> None:
        collection: RegisteredTaskCollection = self.folder.get_tasks()

        self.assertRaises(IndexError, collection.item, 0)

    def test_item_out_of_range(self) -> None:
        collection: RegisteredTaskCollection = self.folder.get_tasks()

        self.assertRaises(IndexError, collection.item, 6)


# TaskVariables  # TODO


class TestRegistrationInfo(unittest.TestCase):
    def test_author(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        registration_info: RegistrationInfo = task_def.registration_info

        expected: str = "Author"
        registration_info.author = expected

        value: str = registration_info.author
        self.assertIsInstance(value, str)
        self.assertEqual(expected, value)

    def test_date(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        registration_info: RegistrationInfo = task_def.registration_info

        expected: Optional[datetime] = datetime.now()
        registration_info.date = expected

        value: Optional[datetime] = registration_info.date
        self.assertIsInstance(value, Optional[datetime])
        self.assertEqual(expected, value)

    def test_date_none(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        registration_info: RegistrationInfo = task_def.registration_info

        expected: Optional[datetime] = None
        registration_info.date = expected

        value: Optional[datetime] = registration_info.date
        self.assertIsInstance(value, Optional[datetime])
        self.assertEqual(expected, value)

    def test_description(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        registration_info: RegistrationInfo = task_def.registration_info

        expected: str = "Description"
        registration_info.description = expected

        value: str = registration_info.description
        self.assertIsInstance(value, str)
        self.assertEqual(expected, value)

    def test_documentation(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        registration_info: RegistrationInfo = task_def.registration_info

        expected: str = "Documentation"
        registration_info.documentation = expected

        value: str = registration_info.documentation
        self.assertIsInstance(value, str)
        self.assertEqual(expected, value)

    def test_security_descriptor(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        registration_info: RegistrationInfo = task_def.registration_info

        expected: Optional[str] = "Security Descriptor"
        registration_info.security_descriptor = expected

        value: Optional[str] = registration_info.security_descriptor
        self.assertIsInstance(value, Optional[str])
        self.assertEqual(expected, value)

    def test_security_descriptor_none(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        registration_info: RegistrationInfo = task_def.registration_info

        expected: Optional[str] = None
        registration_info.security_descriptor = expected

        value: Optional[str] = registration_info.security_descriptor
        self.assertIsInstance(value, Optional[str])
        self.assertEqual(expected, value)

    def test_source(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        registration_info: RegistrationInfo = task_def.registration_info

        expected: str = "Source"
        registration_info.source = expected

        value: str = registration_info.source
        self.assertIsInstance(value, str)
        self.assertEqual(expected, value)

    def test_uri(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        registration_info: RegistrationInfo = task_def.registration_info

        expected: str = "URI"
        registration_info.uri = expected

        value: str = registration_info.uri
        self.assertIsInstance(value, str)
        self.assertEqual(expected, value)

    def test_version(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        registration_info: RegistrationInfo = task_def.registration_info

        expected: str = "Version"
        registration_info.version = expected

        value: str = registration_info.version
        self.assertIsInstance(value, str)
        self.assertEqual(expected, value)

    def test_xml_text(self) -> None:  # TODO
        task_def: TaskDefinition = SERVICE.new_task()
        registration_info: RegistrationInfo = task_def.registration_info

        registration_info.author = "Author"
        registration_info.date = datetime.now()
        registration_info.description = "Description"
        registration_info.documentation = "Documentation"
        registration_info.security_descriptor = "Security Descriptor"
        registration_info.source = "Source"
        registration_info.uri = "Uri"
        registration_info.version = "Version"

        # value: str = registration_info.xml_text
        # pywintypes.com_error: (-2147352567, 'Exception occurred.', (0, None, None, None, 0, -2147467263), None)
        # self.assertIsInstance(value, str)


class TestRepetitionPattern(unittest.TestCase):
    def test_duration(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        trigger: Trigger = task_def.triggers.create(TriggerType.EVENT)
        pattern: RepetitionPattern = trigger.repetition

        expected: relativedelta = relativedelta(seconds=30)
        pattern.duration = expected

        value: relativedelta = pattern.duration
        self.assertIsInstance(value, relativedelta)
        self.assertEqual(expected, value)

    def test_interval(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        trigger: Trigger = task_def.triggers.create(TriggerType.EVENT)
        pattern: RepetitionPattern = trigger.repetition

        expected: relativedelta = relativedelta(seconds=30)
        pattern.interval = expected

        value: relativedelta = pattern.interval
        self.assertIsInstance(value, relativedelta)
        self.assertEqual(expected, value)

    def test_stop_at_duration_end(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        trigger: Trigger = task_def.triggers.create(TriggerType.EVENT)
        pattern: RepetitionPattern = trigger.repetition

        expected: bool = False
        pattern.stop_at_duration_end = expected

        value: bool = pattern.stop_at_duration_end
        self.assertIsInstance(value, bool)
        self.assertEqual(expected, value)


class TestPrincipal(unittest.TestCase):
    def test_display_name(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        principal: Principal = task_def.principal

        expected: str = "Display Name"
        principal.display_name = expected

        value: str = principal.display_name
        self.assertIsInstance(value, str)
        self.assertEqual(expected, value)

    def test_group_id(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        principal: Principal = task_def.principal

        expected: str = "Group ID"
        principal.group_id = expected

        value: str = principal.group_id
        self.assertIsInstance(value, str)
        self.assertEqual(expected, value)

    def test_id(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        principal: Principal = task_def.principal

        expected: str = "ID"
        principal.id = expected

        value: str = principal.id
        self.assertIsInstance(value, str)
        self.assertEqual(expected, value)

    def test_logon_type(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        principal: Principal = task_def.principal

        expected: LogonType = LogonType.INTERACTIVE_TOKEN_OR_PASSWORD
        principal.logon_type = expected

        value: LogonType = principal.logon_type
        self.assertIsInstance(value, LogonType)
        self.assertEqual(expected, value)

    def test_run_level(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        principal: Principal = task_def.principal

        expected: RunLevel = RunLevel.HIGHEST
        principal.run_level = expected

        value: RunLevel = principal.run_level
        self.assertIsInstance(value, RunLevel)
        self.assertEqual(expected, value)

    def test_user_id(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        principal: Principal = task_def.principal

        expected: str = "User ID"
        principal.user_id = expected

        value: str = principal.user_id
        self.assertIsInstance(value, str)
        self.assertEqual(expected, value)


class TestTaskNamedValuePair(unittest.TestCase):
    def test_dunder_len(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        trigger: EventTrigger = cast(EventTrigger, task_def.triggers.create(TriggerType.EVENT))
        pair: TaskNamedValuePair = trigger.value_queries.create("Name1", "Value1")

        self.assertEqual(2, len(pair))

    def test_dunder_getitem(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        trigger: EventTrigger = cast(EventTrigger, task_def.triggers.create(TriggerType.EVENT))
        pair: TaskNamedValuePair = trigger.value_queries.create("Name1", "Value1")

        name: str = pair[0]
        self.assertEqual("Name1", name)

        value: str = pair[1]
        self.assertEqual("Value1", value)

    def test_dunder_getitem_negative(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        trigger: EventTrigger = cast(EventTrigger, task_def.triggers.create(TriggerType.EVENT))
        pair: TaskNamedValuePair = trigger.value_queries.create("Name1", "Value1")

        self.assertRaises(IndexError, pair.__getitem__, -1)

    def test_dunder_getitem_out_of_range(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        trigger: EventTrigger = cast(EventTrigger, task_def.triggers.create(TriggerType.EVENT))
        pair: TaskNamedValuePair = trigger.value_queries.create("Name1", "Value1")

        self.assertRaises(IndexError, pair.__getitem__, 2)

    def test_dunder_iter(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        trigger: EventTrigger = cast(EventTrigger, task_def.triggers.create(TriggerType.EVENT))
        pair: TaskNamedValuePair = trigger.value_queries.create("Name1", "Value1")

        iterator: Iterator[str] = iter(pair)
        self.assertIsInstance(iterator, Iterator)

        for value in pair:
            self.assertIsInstance(value, str)

    def test_tuple(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        trigger: EventTrigger = cast(EventTrigger, task_def.triggers.create(TriggerType.EVENT))
        pair: TaskNamedValuePair = trigger.value_queries.create("Name1", "Value1")

        value: Tuple = tuple(pair)
        self.assertEqual(("Name1", "Value1"), value)

    def test_name(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        trigger: EventTrigger = cast(EventTrigger, task_def.triggers.create(TriggerType.EVENT))
        pair: TaskNamedValuePair = trigger.value_queries.create("Name1", "Value1")

        expected: str = "Name"
        pair.name = expected

        value: str = pair.name
        self.assertIsInstance(value, str)
        self.assertEqual(expected, value)

    def test_value(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        trigger: EventTrigger = cast(EventTrigger, task_def.triggers.create(TriggerType.EVENT))
        pair: TaskNamedValuePair = trigger.value_queries.create("Name1", "Value1")

        expected: str = "Value"
        pair.value = expected

        value: str = pair.value
        self.assertIsInstance(value, str)
        self.assertEqual(expected, value)


class TestTaskNamedValueCollection(unittest.TestCase):
    def test_dunder_len(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        trigger: EventTrigger = cast(EventTrigger, task_def.triggers.create(TriggerType.EVENT))
        collection: TaskNamedValueCollection = trigger.value_queries

        self.assertEqual(0, len(collection))

        collection.create("Name1", "Value1")
        collection.create("Name2", "Value2")
        collection.create("Name3", "Value3")

        self.assertEqual(3, len(collection))

    def test_dunder_getitem(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        trigger: EventTrigger = cast(EventTrigger, task_def.triggers.create(TriggerType.EVENT))
        collection: TaskNamedValueCollection = trigger.value_queries

        collection.create("Name1", "Value1")
        collection.create("Name2", "Value2")
        collection.create("Name3", "Value3")

        pair: TaskNamedValuePair = collection[1]
        self.assertIsInstance(pair, TaskNamedValuePair)
        self.assertEqual("Name1", pair.name)
        self.assertEqual("Value1", pair.value)

    def test_dunder_getitem_zero(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        trigger: EventTrigger = cast(EventTrigger, task_def.triggers.create(TriggerType.EVENT))
        collection: TaskNamedValueCollection = trigger.value_queries

        collection.create("Name1", "Value1")
        collection.create("Name2", "Value2")
        collection.create("Name3", "Value3")

        self.assertRaises(IndexError, collection.__getitem__, 0)

    def test_dunder_getitem_out_of_range(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        trigger: EventTrigger = cast(EventTrigger, task_def.triggers.create(TriggerType.EVENT))
        collection: TaskNamedValueCollection = trigger.value_queries

        collection.create("Name1", "Value1")
        collection.create("Name2", "Value2")
        collection.create("Name3", "Value3")

        self.assertRaises(IndexError, collection.__getitem__, 4)

    def test_dunder_iter(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        trigger: EventTrigger = cast(EventTrigger, task_def.triggers.create(TriggerType.EVENT))
        collection: TaskNamedValueCollection = trigger.value_queries

        collection.create("Name1", "Value1")
        collection.create("Name2", "Value2")
        collection.create("Name3", "Value3")

        iterator: Iterator[TaskNamedValuePair] = iter(collection)
        self.assertIsInstance(iterator, Iterator)

        for member in collection:
            self.assertIsInstance(member, TaskNamedValuePair)

    def test_count(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        trigger: EventTrigger = cast(EventTrigger, task_def.triggers.create(TriggerType.EVENT))
        collection: TaskNamedValueCollection = trigger.value_queries

        self.assertEqual(0, collection.count)

        collection.create("Name1", "Value1")
        collection.create("Name2", "Value2")
        collection.create("Name3", "Value3")

        self.assertEqual(3, collection.count)

    def test_item(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        trigger: EventTrigger = cast(EventTrigger, task_def.triggers.create(TriggerType.EVENT))
        collection: TaskNamedValueCollection = trigger.value_queries

        collection.create("Name1", "Value1")
        collection.create("Name2", "Value2")
        collection.create("Name3", "Value3")

        pair: TaskNamedValuePair = collection.item(1)
        self.assertIsInstance(pair, TaskNamedValuePair)
        self.assertEqual("Name1", pair.name)
        self.assertEqual("Value1", pair.value)

    def test_item_zero(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        trigger: EventTrigger = cast(EventTrigger, task_def.triggers.create(TriggerType.EVENT))
        collection: TaskNamedValueCollection = trigger.value_queries

        collection.create("Name1", "Value1")
        collection.create("Name2", "Value2")
        collection.create("Name3", "Value3")

        self.assertRaises(IndexError, collection.item, 0)

    def test_item_out_of_range(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        trigger: EventTrigger = cast(EventTrigger, task_def.triggers.create(TriggerType.EVENT))
        collection: TaskNamedValueCollection = trigger.value_queries

        collection.create("Name1", "Value1")
        collection.create("Name2", "Value2")
        collection.create("Name3", "Value3")

        self.assertRaises(IndexError, collection.item, 4)

    def test_clear(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        trigger: EventTrigger = cast(EventTrigger, task_def.triggers.create(TriggerType.EVENT))
        collection: TaskNamedValueCollection = trigger.value_queries

        collection.create("Name1", "Value1")
        collection.create("Name2", "Value2")
        collection.create("Name3", "Value3")

        self.assertEqual(3, collection.count)

        collection.clear()
        self.assertEqual(0, collection.count)

    def test_create(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        trigger: EventTrigger = cast(EventTrigger, task_def.triggers.create(TriggerType.EVENT))
        collection: TaskNamedValueCollection = trigger.value_queries

        self.assertEqual(0, collection.count)

        pair: TaskNamedValuePair = collection.create("Name1", "Value1")
        self.assertIsInstance(pair, TaskNamedValuePair)
        self.assertEqual(1, collection.count)

    def test_remove(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        trigger: EventTrigger = cast(EventTrigger, task_def.triggers.create(TriggerType.EVENT))
        collection: TaskNamedValueCollection = trigger.value_queries

        collection.create("Name1", "Value1")
        collection.create("Name2", "Value2")
        collection.create("Name3", "Value3")

        collection.remove(2)
        self.assertEqual(2, collection.count)
        self.assertIsInstance(collection[2], TaskNamedValuePair)


class TestAction(unittest.TestCase):
    def test_id(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        action: Action = task_def.actions.create(ActionType.EXEC)

        expected: str = "ID"
        action.id = expected

        value: str = action.id
        self.assertIsInstance(value, str)
        self.assertEqual(expected, value)

    def test_type(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        action: Action = task_def.actions.create(ActionType.EXEC)

        expected: ActionType = ActionType.EXEC

        value: ActionType = action.type
        self.assertIsInstance(value, ActionType)
        self.assertEqual(expected, value)


class TestActionCollection(unittest.TestCase):
    def test_dunder_len(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        collection: ActionCollection = task_def.actions

        self.assertEqual(0, len(collection))

        collection.create(ActionType.EXEC)
        collection.create(ActionType.COM_HANDLER)
        collection.create(ActionType.SHOW_MESSAGE)

        self.assertEqual(3, len(collection))

    def test_dunder_getitem(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        collection: ActionCollection = task_def.actions

        collection.create(ActionType.EXEC)
        collection.create(ActionType.COM_HANDLER)
        collection.create(ActionType.SHOW_MESSAGE)

        action: Action = collection[1]
        self.assertIsInstance(action, Action)
        self.assertEqual(ActionType.EXEC, action.type)

    def test_dunder_getitem_zero(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        collection: ActionCollection = task_def.actions

        collection.create(ActionType.EXEC)
        collection.create(ActionType.COM_HANDLER)
        collection.create(ActionType.SHOW_MESSAGE)

        self.assertRaises(IndexError, collection.__getitem__, 0)

    def test_dunder_getitem_out_of_range(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        collection: ActionCollection = task_def.actions

        collection.create(ActionType.EXEC)
        collection.create(ActionType.COM_HANDLER)
        collection.create(ActionType.SHOW_MESSAGE)

        self.assertRaises(IndexError, collection.__getitem__, 4)

    def test_dunder_iter(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        collection: ActionCollection = task_def.actions

        collection.create(ActionType.EXEC)
        collection.create(ActionType.COM_HANDLER)
        collection.create(ActionType.SHOW_MESSAGE)

        iterator: Iterator[Action] = iter(collection)
        self.assertIsInstance(iterator, Iterator)

        for member in collection:
            self.assertIsInstance(member, Action)

    def test_context(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        collection: ActionCollection = task_def.actions

        collection.create(ActionType.EXEC)
        collection.create(ActionType.COM_HANDLER)
        collection.create(ActionType.SHOW_MESSAGE)

        expected: str = "Context"
        collection.context = expected

        value: str = collection.context
        self.assertIsInstance(value, str)
        self.assertEqual(expected, value)

    def test_count(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        collection: ActionCollection = task_def.actions

        self.assertEqual(0, collection.count)

        collection.create(ActionType.EXEC)
        collection.create(ActionType.COM_HANDLER)
        collection.create(ActionType.SHOW_MESSAGE)

        self.assertEqual(3, collection.count)

    def test_item(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        collection: ActionCollection = task_def.actions

        collection.create(ActionType.EXEC)
        collection.create(ActionType.COM_HANDLER)
        collection.create(ActionType.SHOW_MESSAGE)

        action: Action = collection.item(1)
        self.assertIsInstance(action, Action)
        self.assertEqual(ActionType.EXEC, action.type)

    def test_dunder_item_zero(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        collection: ActionCollection = task_def.actions

        collection.create(ActionType.EXEC)
        collection.create(ActionType.COM_HANDLER)
        collection.create(ActionType.SHOW_MESSAGE)

        self.assertRaises(IndexError, collection.item, 0)

    def test_dunder_item_out_of_range(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        collection: ActionCollection = task_def.actions

        collection.create(ActionType.EXEC)
        collection.create(ActionType.COM_HANDLER)
        collection.create(ActionType.SHOW_MESSAGE)

        self.assertRaises(IndexError, collection.item, 4)

    def test_clear(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        collection: ActionCollection = task_def.actions

        collection.create(ActionType.EXEC)
        collection.create(ActionType.COM_HANDLER)
        collection.create(ActionType.SHOW_MESSAGE)

        self.assertEqual(3, collection.count)

        collection.clear()
        self.assertEqual(0, collection.count)

    def test_create(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        collection: ActionCollection = task_def.actions

        self.assertEqual(0, collection.count)

        action: Action = collection.create(ActionType.EXEC)
        self.assertIsInstance(action, Action)
        self.assertEqual(1, collection.count)

    def test_remove(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        collection: ActionCollection = task_def.actions

        collection.create(ActionType.EXEC)
        collection.create(ActionType.COM_HANDLER)
        collection.create(ActionType.SHOW_MESSAGE)

        self.assertEqual(3, collection.count)

        collection.remove(2)
        self.assertEqual(2, collection.count)
        self.assertIsInstance(collection[2], ShowMessageAction)


class TestExecAction(unittest.TestCase):
    def test_create(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        action: ExecAction = cast(ExecAction, task_def.actions.create(ActionType.EXEC))

        self.assertIsInstance(action, ExecAction)

    def test_arguments(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        action: ExecAction = cast(ExecAction, task_def.actions.create(ActionType.EXEC))

        expected: str = "Arguments"
        action.arguments = expected

        value: str = action.arguments
        self.assertIsInstance(value, str)
        self.assertEqual(expected, value)

    def test_path(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        action: ExecAction = cast(ExecAction, task_def.actions.create(ActionType.EXEC))

        expected: str = "Path"
        action.path = expected

        value: str = action.path
        self.assertIsInstance(value, str)
        self.assertEqual(expected, value)

    def test_working_directory(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        action: ExecAction = cast(ExecAction, task_def.actions.create(ActionType.EXEC))

        expected: str = "Working Directory"
        action.working_directory = expected

        value: str = action.working_directory
        self.assertIsInstance(value, str)
        self.assertEqual(expected, value)


class TestComHandlerAction(unittest.TestCase):
    def test_create(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        action: ComHandlerAction = cast(
            ComHandlerAction, task_def.actions.create(ActionType.COM_HANDLER)
        )

        self.assertIsInstance(action, ComHandlerAction)

    def test_class_id(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        action: ComHandlerAction = cast(
            ComHandlerAction, task_def.actions.create(ActionType.COM_HANDLER)
        )

        expected: str = "Class ID"
        action.class_id = expected

        value: str = action.class_id
        self.assertIsInstance(value, str)
        self.assertEqual(expected, value)

    def test_data(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        action: ComHandlerAction = cast(
            ComHandlerAction, task_def.actions.create(ActionType.COM_HANDLER)
        )

        expected: str = "Data"
        action.data = expected

        value: str = action.data
        self.assertIsInstance(value, str)
        self.assertEqual(expected, value)


class TestEmailAction(unittest.TestCase):
    def test_create(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        action: EmailAction = cast(EmailAction, task_def.actions.create(ActionType.SEND_EMAIL))

        self.assertIsInstance(action, EmailAction)

    def test_attachments(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        action: EmailAction = cast(EmailAction, task_def.actions.create(ActionType.SEND_EMAIL))

        expected: Optional[Sequence[str]] = ("Attachments",)
        action.attachments = expected

        value: Optional[Sequence[str]] = action.attachments
        self.assertIsInstance(value, Optional[Sequence])
        self.assertEqual(expected, value)

    def test_bcc(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        action: EmailAction = cast(EmailAction, task_def.actions.create(ActionType.SEND_EMAIL))

        expected: str = "BCC"
        action.bcc = expected

        value: str = action.bcc
        self.assertIsInstance(value, str)
        self.assertEqual(expected, value)

    def test_body(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        action: EmailAction = cast(EmailAction, task_def.actions.create(ActionType.SEND_EMAIL))

        expected: str = "Body"
        action.body = expected

        value: str = action.body
        self.assertIsInstance(value, str)
        self.assertEqual(expected, value)

    def test_cc(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        action: EmailAction = cast(EmailAction, task_def.actions.create(ActionType.SEND_EMAIL))

        expected: str = "CC"
        action.cc = expected

        value: str = action.cc
        self.assertIsInstance(value, str)
        self.assertEqual(expected, value)

    def test_from_(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        action: EmailAction = cast(EmailAction, task_def.actions.create(ActionType.SEND_EMAIL))

        expected: str = "From"
        action.from_ = expected

        value: str = action.from_
        self.assertIsInstance(value, str)
        self.assertEqual(expected, value)

    def test_header_fields(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        action: EmailAction = cast(EmailAction, task_def.actions.create(ActionType.SEND_EMAIL))

        value: TaskNamedValueCollection = action.header_fields
        self.assertIsInstance(value, TaskNamedValueCollection)
        self.assertIs(value, action.header_fields)

    def test_reply_to(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        action: EmailAction = cast(EmailAction, task_def.actions.create(ActionType.SEND_EMAIL))

        expected: str = "Reply To"
        action.reply_to = expected

        value: str = action.reply_to
        self.assertIsInstance(value, str)
        self.assertEqual(expected, value)

    def test_server(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        action: EmailAction = cast(EmailAction, task_def.actions.create(ActionType.SEND_EMAIL))

        expected: str = "Server"
        action.server = expected

        value: str = action.server
        self.assertIsInstance(value, str)
        self.assertEqual(expected, value)

    def test_subject(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        action: EmailAction = cast(EmailAction, task_def.actions.create(ActionType.SEND_EMAIL))

        expected: str = "Subject"
        action.subject = expected

        value: str = action.subject
        self.assertIsInstance(value, str)
        self.assertEqual(expected, value)

    def test_to(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        action: EmailAction = cast(EmailAction, task_def.actions.create(ActionType.SEND_EMAIL))

        expected: str = "To"
        action.to = expected

        value: str = action.to
        self.assertIsInstance(value, str)
        self.assertEqual(expected, value)


class TestShowMessageAction(unittest.TestCase):
    def test_create(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        action: ShowMessageAction = cast(
            ShowMessageAction, task_def.actions.create(ActionType.SHOW_MESSAGE)
        )

        self.assertIsInstance(action, ShowMessageAction)

    def test_message_body(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        action: ShowMessageAction = cast(
            ShowMessageAction, task_def.actions.create(ActionType.SHOW_MESSAGE)
        )

        expected: str = "Message Body"
        action.message_body = expected

        value: str = action.message_body
        self.assertIsInstance(value, str)
        self.assertEqual(expected, value)

    def test_title(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        action: ShowMessageAction = cast(
            ShowMessageAction, task_def.actions.create(ActionType.SHOW_MESSAGE)
        )

        expected: str = "Title"
        action.title = expected

        value: str = action.title
        self.assertIsInstance(value, str)
        self.assertEqual(expected, value)


class TestTrigger(unittest.TestCase):
    def test_enabled(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        trigger: Trigger = task_def.triggers.create(TriggerType.EVENT)

        expected: bool = True
        trigger.enabled = expected

        value: bool = trigger.enabled
        self.assertIsInstance(value, bool)
        self.assertEqual(expected, value)

    def test_end_boundary(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        trigger: Trigger = task_def.triggers.create(TriggerType.EVENT)

        expected: Optional[datetime] = datetime.now()
        trigger.end_boundary = expected

        value: Optional[datetime] = trigger.end_boundary
        self.assertIsInstance(value, Optional[datetime])
        self.assertEqual(expected, value)

    def test_execution_time_limit(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        trigger: Trigger = task_def.triggers.create(TriggerType.EVENT)

        expected: relativedelta = relativedelta(days=3)
        trigger.execution_time_limit = expected

        value: relativedelta = trigger.execution_time_limit
        self.assertIsInstance(value, relativedelta)
        self.assertEqual(expected, value)

    def test_id(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        trigger: Trigger = task_def.triggers.create(TriggerType.EVENT)

        expected: str = "ID"
        trigger.id = expected

        value: str = trigger.id
        self.assertIsInstance(value, str)
        self.assertEqual(expected, value)

    def test_repetition(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        trigger: Trigger = task_def.triggers.create(TriggerType.EVENT)

        value: RepetitionPattern = trigger.repetition
        self.assertIsInstance(value, RepetitionPattern)
        self.assertIs(value, trigger.repetition)

    def test_start_boundary(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        trigger: Trigger = task_def.triggers.create(TriggerType.EVENT)

        expected: Optional[datetime] = datetime.now()
        trigger.start_boundary = expected

        value: Optional[datetime] = trigger.start_boundary
        self.assertIsInstance(value, Optional[datetime])
        self.assertEqual(expected, value)

    def test_type(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        trigger: Trigger = task_def.triggers.create(TriggerType.EVENT)

        expected: TriggerType = TriggerType.EVENT

        value: TriggerType = trigger.type
        self.assertIsInstance(value, TriggerType)
        self.assertEqual(expected, value)


class TestTriggerCollection(unittest.TestCase):
    def test_dunder_len(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        collection: TriggerCollection = task_def.triggers

        self.assertEqual(0, len(collection))

        collection.create(TriggerType.EVENT)
        collection.create(TriggerType.TIME)
        collection.create(TriggerType.DAILY)

        self.assertEqual(3, len(collection))

    def test_dunder_getitem(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        collection: TriggerCollection = task_def.triggers

        collection.create(TriggerType.EVENT)
        collection.create(TriggerType.TIME)
        collection.create(TriggerType.DAILY)

        trigger: Trigger = collection[1]
        self.assertIsInstance(trigger, Trigger)
        self.assertEqual(TriggerType.EVENT, trigger.type)

    def test_dunder_getitem_zero(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        collection: TriggerCollection = task_def.triggers

        collection.create(TriggerType.EVENT)
        collection.create(TriggerType.TIME)
        collection.create(TriggerType.DAILY)

        self.assertRaises(IndexError, collection.__getitem__, 0)

    def test_dunder_getitem_out_of_range(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        collection: TriggerCollection = task_def.triggers

        collection.create(TriggerType.EVENT)
        collection.create(TriggerType.TIME)
        collection.create(TriggerType.DAILY)

        self.assertRaises(IndexError, collection.__getitem__, 4)

    def test_dunder_iter(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        collection: TriggerCollection = task_def.triggers

        collection.create(TriggerType.EVENT)
        collection.create(TriggerType.TIME)
        collection.create(TriggerType.DAILY)

        iterator: Iterator[Trigger] = iter(collection)
        self.assertIsInstance(iterator, Iterator)

        for member in collection:
            self.assertIsInstance(member, Trigger)

    def test_count(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        collection: TriggerCollection = task_def.triggers

        self.assertEqual(0, collection.count)

        collection.create(TriggerType.EVENT)
        collection.create(TriggerType.TIME)
        collection.create(TriggerType.DAILY)

        self.assertEqual(3, collection.count)

    def test_item(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        collection: TriggerCollection = task_def.triggers

        collection.create(TriggerType.EVENT)
        collection.create(TriggerType.TIME)
        collection.create(TriggerType.DAILY)

        trigger: Trigger = collection.item(1)
        self.assertIsInstance(trigger, Trigger)
        self.assertEqual(TriggerType.EVENT, trigger.type)

    def test_item_zero(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        collection: TriggerCollection = task_def.triggers

        collection.create(TriggerType.EVENT)
        collection.create(TriggerType.TIME)
        collection.create(TriggerType.DAILY)

        self.assertRaises(IndexError, collection.item, 0)

    def test_item_out_of_range(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        collection: TriggerCollection = task_def.triggers

        collection.create(TriggerType.EVENT)
        collection.create(TriggerType.TIME)
        collection.create(TriggerType.DAILY)

        self.assertRaises(IndexError, collection.item, 4)

    def test_clear(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        collection: TriggerCollection = task_def.triggers

        collection.create(TriggerType.EVENT)
        collection.create(TriggerType.TIME)
        collection.create(TriggerType.DAILY)

        self.assertEqual(3, collection.count)

        collection.clear()
        self.assertEqual(0, collection.count)

    def test_create(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        collection: TriggerCollection = task_def.triggers

        self.assertEqual(0, collection.count)

        trigger: Trigger = collection.create(TriggerType.EVENT)
        self.assertIsInstance(trigger, Trigger)
        self.assertEqual(1, collection.count)

    def test_remove(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        collection: TriggerCollection = task_def.triggers

        collection.create(TriggerType.EVENT)
        collection.create(TriggerType.TIME)
        collection.create(TriggerType.DAILY)

        self.assertEqual(3, collection.count)

        collection.remove(2)
        self.assertEqual(2, collection.count)
        self.assertIsInstance(collection[2], DailyTrigger)


class TestEventTrigger(unittest.TestCase):
    def test_create(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        trigger: EventTrigger = cast(EventTrigger, task_def.triggers.create(TriggerType.EVENT))

        self.assertIsInstance(trigger, EventTrigger)

    def test_delay(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        trigger: EventTrigger = cast(EventTrigger, task_def.triggers.create(TriggerType.EVENT))

        expected: relativedelta = relativedelta(minutes=30)
        trigger.delay = expected

        value: relativedelta = trigger.delay
        self.assertIsInstance(value, relativedelta)
        self.assertEqual(expected, value)

    def test_subscription(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        trigger: EventTrigger = cast(EventTrigger, task_def.triggers.create(TriggerType.EVENT))

        expected: str = "Subscription"
        trigger.subscription = expected

        value: str = trigger.subscription
        self.assertIsInstance(value, str)
        self.assertEqual(expected, value)

    def test_value_queries(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        trigger: EventTrigger = cast(EventTrigger, task_def.triggers.create(TriggerType.EVENT))

        value: TaskNamedValueCollection = trigger.value_queries
        self.assertIsInstance(value, TaskNamedValueCollection)
        self.assertIs(value, trigger.value_queries)


class TestTimeTrigger(unittest.TestCase):
    def test_create(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        trigger: TimeTrigger = cast(TimeTrigger, task_def.triggers.create(TriggerType.TIME))

        self.assertIsInstance(trigger, TimeTrigger)

    def test_random_delay(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        trigger: TimeTrigger = cast(TimeTrigger, task_def.triggers.create(TriggerType.TIME))

        expected: relativedelta = relativedelta(minutes=30)
        trigger.random_delay = expected

        value: relativedelta = trigger.random_delay
        self.assertIsInstance(value, relativedelta)
        self.assertEqual(expected, value)


class TestDailyTrigger(unittest.TestCase):
    def test_create(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        trigger: DailyTrigger = cast(DailyTrigger, task_def.triggers.create(TriggerType.DAILY))

        self.assertIsInstance(trigger, DailyTrigger)

    def test_days_interval(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        trigger: DailyTrigger = cast(DailyTrigger, task_def.triggers.create(TriggerType.DAILY))

        expected: int = 3
        trigger.days_interval = expected

        value: int = trigger.days_interval
        self.assertIsInstance(value, int)
        self.assertEqual(expected, value)


class TestWeeklyTrigger(unittest.TestCase):
    def test_create(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        trigger: WeeklyTrigger = cast(WeeklyTrigger, task_def.triggers.create(TriggerType.WEEKLY))

        self.assertIsInstance(trigger, WeeklyTrigger)

    def test_days_of_week(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        trigger: WeeklyTrigger = cast(WeeklyTrigger, task_def.triggers.create(TriggerType.WEEKLY))

        expected: DaysOfWeek = DaysOfWeek.THURSDAY
        trigger.days_of_week = expected

        value: DaysOfWeek = trigger.days_of_week
        self.assertIsInstance(value, DaysOfWeek)
        self.assertEqual(expected, value)

    def test_weeks_interval(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        trigger: WeeklyTrigger = cast(WeeklyTrigger, task_def.triggers.create(TriggerType.WEEKLY))

        expected: int = 3
        trigger.weeks_interval = expected

        value: int = trigger.weeks_interval
        self.assertIsInstance(value, int)
        self.assertEqual(expected, value)


class TestMonthlyTrigger(unittest.TestCase):
    def test_create(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        trigger: MonthlyTrigger = cast(
            MonthlyTrigger, task_def.triggers.create(TriggerType.MONTHLY)
        )

        self.assertIsInstance(trigger, MonthlyTrigger)

    def test_days_of_month(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        trigger: MonthlyTrigger = cast(
            MonthlyTrigger, task_def.triggers.create(TriggerType.MONTHLY)
        )

        expected: DaysOfMonth = DaysOfMonth.THIRTEENTH
        trigger.days_of_month = expected

        value: DaysOfMonth = trigger.days_of_month
        self.assertIsInstance(value, DaysOfMonth)
        self.assertEqual(expected, value)

    def test_months_of_year(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        trigger: MonthlyTrigger = cast(
            MonthlyTrigger, task_def.triggers.create(TriggerType.MONTHLY)
        )

        expected: MonthsOfYear = MonthsOfYear.FEBRUARY
        trigger.months_of_year = expected

        value: MonthsOfYear = trigger.months_of_year
        self.assertIsInstance(value, MonthsOfYear)
        self.assertEqual(expected, value)

    def test_run_on_last_day_of_month(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        trigger: MonthlyTrigger = cast(
            MonthlyTrigger, task_def.triggers.create(TriggerType.MONTHLY)
        )

        expected: bool = True
        trigger.run_on_last_day_of_month = expected

        value: bool = trigger.run_on_last_day_of_month
        self.assertIsInstance(value, bool)
        self.assertEqual(expected, value)


class TestMonthlyDOWTrigger(unittest.TestCase):
    def test_create(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        trigger: MonthlyDOWTrigger = cast(
            MonthlyDOWTrigger, task_def.triggers.create(TriggerType.MONTHLY_DOW)
        )

        self.assertIsInstance(trigger, MonthlyDOWTrigger)

    def test_days_of_week(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        trigger: MonthlyDOWTrigger = cast(
            MonthlyDOWTrigger, task_def.triggers.create(TriggerType.MONTHLY)
        )

        expected: DaysOfWeek = DaysOfWeek.TUESDAY
        trigger.days_of_week = expected

        value: DaysOfWeek = trigger.days_of_week
        self.assertIsInstance(value, DaysOfWeek)
        self.assertEqual(expected, value)

    def test_months_of_year(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        trigger: MonthlyDOWTrigger = cast(
            MonthlyDOWTrigger, task_def.triggers.create(TriggerType.MONTHLY)
        )

        expected: MonthsOfYear = MonthsOfYear.AUGUST
        trigger.months_of_year = expected

        value: MonthsOfYear = trigger.months_of_year
        self.assertIsInstance(value, MonthsOfYear)
        self.assertEqual(expected, value)

    def test_run_on_last_week_of_month(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        trigger: MonthlyDOWTrigger = cast(
            MonthlyDOWTrigger, task_def.triggers.create(TriggerType.MONTHLY)
        )

        expected: bool = True
        trigger.run_on_last_week_of_month = expected

        value: bool = trigger.run_on_last_week_of_month
        self.assertIsInstance(value, bool)
        self.assertEqual(expected, value)

    def test_weeks_of_month(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        trigger: MonthlyDOWTrigger = cast(
            MonthlyDOWTrigger, task_def.triggers.create(TriggerType.MONTHLY)
        )

        expected: WeeksOfMonth = WeeksOfMonth.SECOND
        trigger.weeks_of_month = expected

        value: WeeksOfMonth = trigger.weeks_of_month
        self.assertIsInstance(value, WeeksOfMonth)
        self.assertEqual(expected, value)


class TestIdleTrigger(unittest.TestCase):
    def test_create(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        trigger: IdleTrigger = cast(IdleTrigger, task_def.triggers.create(TriggerType.IDLE))

        self.assertIsInstance(trigger, IdleTrigger)


class TestRegistrationTrigger(unittest.TestCase):
    def test_create(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        trigger: RegistrationTrigger = cast(
            RegistrationTrigger, task_def.triggers.create(TriggerType.REGISTRATION)
        )

        self.assertIsInstance(trigger, RegistrationTrigger)

    def test_delay(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        trigger: RegistrationTrigger = cast(
            RegistrationTrigger, task_def.triggers.create(TriggerType.REGISTRATION)
        )

        expected: relativedelta = relativedelta(hours=5)
        trigger.delay = expected

        value: relativedelta = trigger.delay
        self.assertIsInstance(value, relativedelta)
        self.assertEqual(expected, value)


class TestBootTrigger(unittest.TestCase):
    def test_create(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        trigger: BootTrigger = cast(BootTrigger, task_def.triggers.create(TriggerType.BOOT))

        self.assertIsInstance(trigger, BootTrigger)

    def test_delay(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        trigger: BootTrigger = cast(BootTrigger, task_def.triggers.create(TriggerType.BOOT))

        expected: relativedelta = relativedelta(hours=5)
        trigger.delay = expected

        value: relativedelta = trigger.delay
        self.assertIsInstance(value, relativedelta)
        self.assertEqual(expected, value)


class TestLogonTrigger(unittest.TestCase):
    def test_create(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        trigger: LogonTrigger = cast(LogonTrigger, task_def.triggers.create(TriggerType.LOGON))

        self.assertIsInstance(trigger, LogonTrigger)

    def test_delay(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        trigger: LogonTrigger = cast(LogonTrigger, task_def.triggers.create(TriggerType.LOGON))

        expected: relativedelta = relativedelta(hours=5)
        trigger.delay = expected

        value: relativedelta = trigger.delay
        self.assertIsInstance(value, relativedelta)
        self.assertEqual(expected, value)

    def test_user_id(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        trigger: LogonTrigger = cast(LogonTrigger, task_def.triggers.create(TriggerType.LOGON))

        expected: str = "User ID"
        trigger.user_id = expected

        value: str = trigger.user_id
        self.assertIsInstance(value, str)
        self.assertEqual(expected, value)


class TestSessionStateChangeTrigger(unittest.TestCase):
    def test_create(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        trigger: SessionStateChangeTrigger = cast(
            SessionStateChangeTrigger, task_def.triggers.create(TriggerType.SESSION_STATE_CHANGE)
        )

        self.assertIsInstance(trigger, SessionStateChangeTrigger)

    def test_delay(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        trigger: SessionStateChangeTrigger = cast(
            SessionStateChangeTrigger, task_def.triggers.create(TriggerType.SESSION_STATE_CHANGE)
        )

        expected: relativedelta = relativedelta(hours=5)
        trigger.delay = expected

        value: relativedelta = trigger.delay
        self.assertIsInstance(value, relativedelta)
        self.assertEqual(expected, value)

    def test_state_change(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        trigger: SessionStateChangeTrigger = cast(
            SessionStateChangeTrigger, task_def.triggers.create(TriggerType.SESSION_STATE_CHANGE)
        )

        expected: SessionStateChangeType = SessionStateChangeType.SESSION_LOCK
        trigger.state_change = expected

        value: SessionStateChangeType = trigger.state_change
        self.assertIsInstance(value, SessionStateChangeType)
        self.assertEqual(expected, value)

    def test_user_id(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        trigger: SessionStateChangeTrigger = cast(
            SessionStateChangeTrigger, task_def.triggers.create(TriggerType.SESSION_STATE_CHANGE)
        )

        expected: str = "User ID"
        trigger.user_id = expected

        value: str = trigger.user_id
        self.assertIsInstance(value, str)
        self.assertEqual(expected, value)


class TestTaskSettings(unittest.TestCase):
    def test_allow_demand_start(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        settings: TaskSettings = task_def.settings

        expected: bool = True
        settings.allow_demand_start = expected

        value: bool = settings.allow_demand_start
        self.assertIsInstance(value, bool)
        self.assertEqual(expected, value)

    def test_allow_hard_terminate(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        settings: TaskSettings = task_def.settings

        expected: bool = True
        settings.allow_hard_terminate = expected

        value: bool = settings.allow_hard_terminate
        self.assertIsInstance(value, bool)
        self.assertEqual(expected, value)

    def test_compatibility(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        settings: TaskSettings = task_def.settings

        expected: Compatibility = Compatibility.V1
        settings.compatibility = expected

        value: Compatibility = settings.compatibility
        self.assertIsInstance(value, Compatibility)
        self.assertEqual(expected, value)

    def test_delete_expired_task_after(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        settings: TaskSettings = task_def.settings

        expected: Optional[datetime] = datetime.now()
        settings.delete_expired_task_after = expected

        value: Optional[datetime] = settings.delete_expired_task_after
        self.assertIsInstance(value, Optional[datetime])
        self.assertEqual(expected, value)

    def test_disallow_start_if_on_batteries(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        settings: TaskSettings = task_def.settings

        expected: bool = True
        settings.disallow_start_if_on_batteries = expected

        value: bool = settings.disallow_start_if_on_batteries
        self.assertIsInstance(value, bool)
        self.assertEqual(expected, value)

    def test_enabled(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        settings: TaskSettings = task_def.settings

        expected: bool = True
        settings.enabled = expected

        value: bool = settings.enabled
        self.assertIsInstance(value, bool)
        self.assertEqual(expected, value)

    def test_execution_time_limit(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        settings: TaskSettings = task_def.settings

        expected: Optional[relativedelta] = relativedelta(days=10)
        settings.execution_time_limit = expected

        value: Optional[relativedelta] = settings.execution_time_limit
        self.assertIsInstance(value, Optional[relativedelta])
        self.assertEqual(expected, value)

    def test_execution_time_limit_none(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        settings: TaskSettings = task_def.settings

        expected: Optional[relativedelta] = None
        settings.execution_time_limit = expected

        value: Optional[relativedelta] = settings.execution_time_limit
        self.assertIsInstance(value, Optional[relativedelta])
        self.assertEqual(expected, value)

    def test_hidden(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        settings: TaskSettings = task_def.settings

        expected: bool = False
        settings.hidden = expected

        value: bool = settings.hidden
        self.assertIsInstance(value, bool)
        self.assertEqual(expected, value)

    def test_idle_settings(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        settings: TaskSettings = task_def.settings

        value: IdleSettings = settings.idle_settings
        self.assertIsInstance(value, IdleSettings)
        self.assertIs(value, settings.idle_settings)

    def test_multiple_instances(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        settings: TaskSettings = task_def.settings

        expected: InstancesPolicy = InstancesPolicy.IGNORE_NEW
        settings.multiple_instances = expected

        value: InstancesPolicy = settings.multiple_instances
        self.assertIsInstance(value, InstancesPolicy)
        self.assertEqual(expected, value)

    def test_network_settings(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        settings: TaskSettings = task_def.settings

        value: NetworkSettings = settings.network_settings
        self.assertIsInstance(value, NetworkSettings)
        self.assertIs(value, settings.network_settings)

    def test_priority(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        settings: TaskSettings = task_def.settings

        expected: int = 8
        settings.priority = expected

        value: int = settings.priority
        self.assertIsInstance(value, int)
        self.assertEqual(expected, value)

    def test_restart_count(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        settings: TaskSettings = task_def.settings

        expected: int = -1
        settings.restart_count = expected

        value: int = settings.restart_count
        self.assertIsInstance(value, int)
        self.assertEqual(expected, value)

    def test_restart_interval(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        settings: TaskSettings = task_def.settings

        expected: Optional[relativedelta] = relativedelta(hours=1)
        settings.restart_interval = expected

        value: Optional[relativedelta] = settings.restart_interval
        self.assertIsInstance(value, Optional[relativedelta])
        self.assertEqual(expected, value)

    def test_run_only_if_idle(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        settings: TaskSettings = task_def.settings

        expected: bool = True
        settings.run_only_if_idle = expected

        value: bool = settings.run_only_if_idle
        self.assertIsInstance(value, bool)
        self.assertEqual(expected, value)

    def test_run_only_if_network_available(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        settings: TaskSettings = task_def.settings

        expected: bool = True
        settings.run_only_if_network_available = expected

        value: bool = settings.run_only_if_network_available
        self.assertIsInstance(value, bool)
        self.assertEqual(expected, value)

    def test_start_when_available(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        settings: TaskSettings = task_def.settings

        expected: bool = True
        settings.start_when_available = expected

        value: bool = settings.start_when_available
        self.assertIsInstance(value, bool)
        self.assertEqual(expected, value)

    def test_stop_if_going_on_batteries(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        settings: TaskSettings = task_def.settings

        expected: bool = True
        settings.stop_if_going_on_batteries = expected

        value: bool = settings.stop_if_going_on_batteries
        self.assertIsInstance(value, bool)
        self.assertEqual(expected, value)

    def test_wake_to_run(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        settings: TaskSettings = task_def.settings

        expected: bool = True
        settings.wake_to_run = expected

        value: bool = settings.wake_to_run
        self.assertIsInstance(value, bool)
        self.assertEqual(expected, value)

    def test_xml_text(self) -> None:  # TODO
        task_def: TaskDefinition = SERVICE.new_task()
        settings: TaskSettings = task_def.settings

        settings.enabled = True

        # value: str = settings.xml_text
        # self.assertIsInstance(value, str)


class TestIdleSettings(unittest.TestCase):
    def test_restart_on_idle(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        task_settings: TaskSettings = task_def.settings
        settings: IdleSettings = task_settings.idle_settings

        expected: bool = True
        settings.restart_on_idle = expected

        value: bool = settings.restart_on_idle
        self.assertIsInstance(value, bool)
        self.assertEqual(expected, value)

    def test_stop_on_idle_end(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        task_settings: TaskSettings = task_def.settings
        settings: IdleSettings = task_settings.idle_settings

        expected: bool = True
        settings.stop_on_idle_end = expected

        value: bool = settings.stop_on_idle_end
        self.assertIsInstance(value, bool)
        self.assertEqual(expected, value)


class TestNetworkSettings(unittest.TestCase):
    def test_id(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        task_settings: TaskSettings = task_def.settings
        settings: NetworkSettings = task_settings.network_settings

        expected: str = "ID"
        settings.id = expected

        value: str = settings.id
        self.assertIsInstance(value, str)
        self.assertEqual(expected, value)

    def test_name(self) -> None:
        task_def: TaskDefinition = SERVICE.new_task()
        task_settings: TaskSettings = task_def.settings
        settings: NetworkSettings = task_settings.network_settings

        expected: str = "Name"
        settings.name = expected

        value: str = settings.name
        self.assertIsInstance(value, str)
        self.assertEqual(expected, value)


class TestFromDurationStr(unittest.TestCase):
    def test_empty(self) -> None:
        string: str = ""
        expected: Optional[relativedelta] = None

        result: Optional[relativedelta] = from_duration_str(string)
        self.assertIsInstance(result, Optional[relativedelta])
        self.assertEqual(expected, result)

    def test_zeros(self) -> None:
        string: str = "P0Y0M0DT0H0M0S"
        expected: Optional[relativedelta] = None

        result: Optional[relativedelta] = from_duration_str(string)
        self.assertIsInstance(result, Optional[relativedelta])
        self.assertEqual(expected, result)

    def test_ones(self) -> None:
        string: str = "P1Y1M1DT1H1M1S"
        expected: Optional[relativedelta] = relativedelta(
            years=1, months=1, days=1, hours=1, minutes=1, seconds=1
        )

        result: Optional[relativedelta] = from_duration_str(string)
        self.assertIsInstance(result, Optional[relativedelta])
        self.assertEqual(expected, result)

    def test_years(self) -> None:
        string: str = "P1YT"
        expected: Optional[relativedelta] = relativedelta(years=1)

        result: Optional[relativedelta] = from_duration_str(string)
        self.assertIsInstance(result, Optional[relativedelta])
        self.assertEqual(expected, result)

    def test_months(self) -> None:
        string: str = "P1MT"
        expected: Optional[relativedelta] = relativedelta(months=1)

        result: Optional[relativedelta] = from_duration_str(string)
        self.assertIsInstance(result, Optional[relativedelta])
        self.assertEqual(expected, result)

    def test_days(self) -> None:
        string: str = "P1DT"
        expected: Optional[relativedelta] = relativedelta(days=1)

        result: Optional[relativedelta] = from_duration_str(string)
        self.assertIsInstance(result, Optional[relativedelta])
        self.assertEqual(expected, result)

    def test_hours(self) -> None:
        string: str = "PT1H"
        expected: Optional[relativedelta] = relativedelta(hours=1)

        result: Optional[relativedelta] = from_duration_str(string)
        self.assertIsInstance(result, Optional[relativedelta])
        self.assertEqual(expected, result)

    def test_minutes(self) -> None:
        string: str = "PT1M"
        expected: Optional[relativedelta] = relativedelta(minutes=1)

        result: Optional[relativedelta] = from_duration_str(string)
        self.assertIsInstance(result, Optional[relativedelta])
        self.assertEqual(expected, result)

    def test_seconds(self) -> None:
        string: str = "PT1S"
        expected: Optional[relativedelta] = relativedelta(seconds=1)

        result: Optional[relativedelta] = from_duration_str(string)
        self.assertIsInstance(result, Optional[relativedelta])
        self.assertEqual(expected, result)


class TestToDurationStr(unittest.TestCase):
    def test_none(self) -> None:
        value: Optional[relativedelta] = None
        expected: str = "PT0S"

        result: str = to_duration_str(value)
        self.assertIsInstance(result, str)
        self.assertEqual(expected, result)

    def test_zeros(self) -> None:
        value: Optional[relativedelta] = relativedelta(
            years=0, months=0, days=0, hours=0, minutes=0, seconds=0
        )
        expected: str = "PT0S"

        result: str = to_duration_str(value)
        self.assertIsInstance(result, str)
        self.assertEqual(expected, result)

    def test_ones(self) -> None:
        value: Optional[relativedelta] = relativedelta(
            years=1, months=1, days=1, hours=1, minutes=1, seconds=1
        )
        expected: str = "P1Y1M1DT1H1M1S"

        result: str = to_duration_str(value)
        self.assertIsInstance(result, str)
        self.assertEqual(expected, result)

    def test_years(self) -> None:
        value: Optional[relativedelta] = relativedelta(years=1)
        expected: str = "P1YT"

        result: str = to_duration_str(value)
        self.assertIsInstance(result, str)
        self.assertEqual(expected, result)

    def test_months(self) -> None:
        value: Optional[relativedelta] = relativedelta(months=1)
        expected: str = "P1MT"

        result: str = to_duration_str(value)
        self.assertIsInstance(result, str)
        self.assertEqual(expected, result)

    def test_days(self) -> None:
        value: Optional[relativedelta] = relativedelta(days=1)
        expected: str = "P1DT"

        result: str = to_duration_str(value)
        self.assertIsInstance(result, str)
        self.assertEqual(expected, result)

    def test_hours(self) -> None:
        value: Optional[relativedelta] = relativedelta(hours=1)
        expected: str = "PT1H"

        result: str = to_duration_str(value)
        self.assertIsInstance(result, str)
        self.assertEqual(expected, result)

    def test_minutes(self) -> None:
        value: Optional[relativedelta] = relativedelta(minutes=1)
        expected: str = "PT1M"

        result: str = to_duration_str(value)
        self.assertIsInstance(result, str)
        self.assertEqual(expected, result)

    def test_seconds(self) -> None:
        value: Optional[relativedelta] = relativedelta(seconds=1)
        expected: str = "PT1S"

        result: str = to_duration_str(value)
        self.assertIsInstance(result, str)
        self.assertEqual(expected, result)


if __name__ == "__main__":
    unittest.main()
