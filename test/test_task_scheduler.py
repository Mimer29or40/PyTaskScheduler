"""Tests for task_scheduler.py."""

from __future__ import annotations

import contextlib
from collections.abc import Iterator
from collections.abc import Mapping
from collections.abc import Sequence
from datetime import datetime
from datetime import timedelta
from pathlib import Path
from typing import TYPE_CHECKING
from typing import Any
from typing import Final

import pytest
from dateutil.relativedelta import relativedelta

from task_scheduler import Action
from task_scheduler import ActionCollection
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
from task_scheduler import RunFlags
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
from task_scheduler import TaskFolderExistsError
from task_scheduler import TaskFolderNotFoundError
from task_scheduler import TaskNamedValueCollection
from task_scheduler import TaskNamedValuePair
from task_scheduler import TaskNotFoundError
from task_scheduler import TaskService
from task_scheduler import TaskSettings
from task_scheduler import TimeTrigger
from task_scheduler import Trigger
from task_scheduler import TriggerCollection
from task_scheduler import WeeklyTrigger
from task_scheduler import WeeksOfMonth
from task_scheduler import from_duration_str
from task_scheduler import to_duration_str

if TYPE_CHECKING:
    from collections.abc import Generator

SERVICE: Final[TaskService] = TaskService()
SERVICE.connect()

ROOT = SERVICE.get_folder("\\")


class TestTaskService:
    """Tests for TaskService."""

    def test_same_instance(self) -> None:
        """Test to check that only one instance of TaskService is created."""
        new_service = TaskService()

        assert SERVICE is new_service

    def test_connected(self) -> None:
        """Test for TaskService.connected."""
        connected: bool = SERVICE.connected
        assert isinstance(connected, bool)

    def test_connected_domain(self) -> None:
        """Test for TaskService.connected_domain."""
        connected_domain: str = SERVICE.connected_domain
        assert isinstance(connected_domain, str)

    def test_connected_user(self) -> None:
        """Test for TaskService.connected_user."""
        connected_user: str = SERVICE.connected_user
        assert isinstance(connected_user, str)

    def test_highest_version(self) -> None:
        """Test for TaskService.highest_version."""
        highest_version: int = SERVICE.highest_version
        assert isinstance(highest_version, int)

    def test_target_server(self) -> None:
        """Test for TaskService.target_server."""
        target_server: str = SERVICE.target_server
        assert isinstance(target_server, str)

    def test_connect(self) -> None:
        """Test for TaskService.connect()."""
        # TODO(Ryan): test_connect

    def test_get_folder(self) -> None:
        """Test for TaskService.get_folder()."""
        folder: TaskFolder = SERVICE.get_folder("\\")
        assert isinstance(folder, TaskFolder)

    def test_get_folder_not_exists(self) -> None:
        """Test for TaskService.get_folder() with a folder that does not exist."""
        with pytest.raises(TaskFolderNotFoundError):
            SERVICE.get_folder("\\FolderNotExists")

    def test_get_running_task(self) -> None:
        """Test for TaskService.get_running_tasks()."""
        tasks: RunningTaskCollection = SERVICE.get_running_tasks()
        assert isinstance(tasks, RunningTaskCollection)

        hidden_tasks: RunningTaskCollection = SERVICE.get_running_tasks(True)
        assert isinstance(hidden_tasks, RunningTaskCollection)

    def test_new_task(self) -> None:
        """Test for TaskService.new_task()."""
        new_task: TaskDefinition = SERVICE.new_task()
        assert isinstance(new_task, TaskDefinition)


class TestTaskFolder:
    """Tests for TaskFolder."""

    def test_dunder_eq(self) -> None:
        """Test for TaskService.__eq__()."""
        folder: TaskFolder = ROOT.create_folder("TestEq")
        try:
            other: TaskFolder = ROOT.get_folder(folder.path)
            assert folder == other
        finally:
            ROOT.delete_folder(folder.path)

    def test_name(self) -> None:
        """Test for TaskFolder.name."""
        name: str = ROOT.name
        assert isinstance(name, str)

    def test_path(self) -> None:
        """Test for TaskFolder.path."""
        path: str = ROOT.path
        assert isinstance(path, str)

    @pytest.mark.parametrize("absolute", [False, True])
    def test_create_folder(self, absolute: bool) -> None:
        """Test for TaskService.create_folder()."""
        name: str = "TestCreateFolder"

        folder_path: str = f"\\{name}" if absolute else name
        folder: TaskFolder = ROOT.create_folder(folder_path)
        try:
            assert isinstance(folder, TaskFolder)
            assert folder.name == name
            assert folder.path == f"\\{name}"
        finally:
            with contextlib.suppress(TaskFolderNotFoundError):
                ROOT.delete_folder(folder_path)

    @pytest.mark.parametrize("absolute", [False, True])
    def test_create_folder_nested(self, absolute: bool) -> None:
        """Test for TaskService.create_folder() with a nested folder."""
        parent_name: str = "TestCreateFolder"
        name: str = "Folder"

        folder_path: str = f"\\{parent_name}\\{name}" if absolute else f"{parent_name}\\{name}"
        folder: TaskFolder = ROOT.create_folder(folder_path)
        try:
            assert isinstance(folder, TaskFolder)
            assert folder.name == name
            assert folder.path == f"\\{parent_name}\\{name}"
        finally:
            with contextlib.suppress(TaskFolderNotFoundError):
                ROOT.delete_folder(folder_path)
            with contextlib.suppress(TaskFolderNotFoundError):
                ROOT.delete_folder(parent_name)

    def test_create_folder_exists(self) -> None:
        """Test for TaskService.create_folder() with a folder that already exists."""
        name: str = "TestCreateFolderExists"

        with contextlib.suppress(TaskFolderExistsError):
            ROOT.create_folder(name)
        try:
            with pytest.raises(TaskFolderExistsError):
                ROOT.create_folder(name)
        finally:
            with contextlib.suppress(TaskFolderNotFoundError):
                ROOT.delete_folder(name)

    def test_create_folder_security_descriptor(self) -> None:
        """Test for TaskService.create_folder() with a security descriptor."""
        # TODO(Ryan): test_create_folder_security_descriptor

    @pytest.mark.parametrize("absolute", [False, True])
    def test_delete_folder(self, absolute: bool) -> None:
        """Test for TaskService.delete_folder()."""
        name: str = "TestDeleteFolder"

        folder_path: str = f"\\{name}" if absolute else name
        with contextlib.suppress(TaskFolderExistsError):
            ROOT.create_folder(folder_path)
        try:
            assert ROOT.get_folder(folder_path) is not None
            ROOT.delete_folder(folder_path)
            with pytest.raises(TaskFolderNotFoundError):
                ROOT.get_folder(folder_path)
        finally:
            with contextlib.suppress(TaskFolderNotFoundError):
                ROOT.delete_folder(folder_path)

    @pytest.mark.parametrize("absolute", [False, True])
    def test_delete_folder_nested(self, absolute: bool) -> None:
        """Test for TaskService.delete_folder() with a nested folder."""
        parent_name: str = "TestDeleteFolder"
        name: str = "Folder"

        folder_path: str = f"\\{parent_name}\\{name}" if absolute else f"{parent_name}\\{name}"
        with contextlib.suppress(TaskFolderExistsError):
            ROOT.create_folder(folder_path)
        try:
            assert ROOT.get_folder(folder_path) is not None
            ROOT.delete_folder(folder_path)
            with pytest.raises(TaskFolderNotFoundError):
                ROOT.get_folder(folder_path)
        finally:
            with contextlib.suppress(TaskFolderNotFoundError):
                ROOT.delete_folder(folder_path)
            with contextlib.suppress(TaskFolderNotFoundError):
                ROOT.delete_folder(parent_name)

    def test_delete_folder_not_exists(self) -> None:
        """Test for TaskService.delete_folder() with a folder that does not exist exists."""
        name: str = "TestDeleteFolderNotExists"

        with contextlib.suppress(TaskFolderNotFoundError):
            ROOT.delete_folder(name)

    def test_delete_task(self) -> None:
        """Test for TaskFolder.delete_task()."""
        folder_name: str = "TestDeleteTask"
        folder: TaskFolder
        try:
            folder = ROOT.create_folder(folder_name)
        except TaskFolderExistsError:
            folder = ROOT.get_folder(folder_name)

        # TODO(Ryan): Implement
        task_def: TaskDefinition = SERVICE.new_task()
        task_def.registration_info.description = "Test Task"
        task_def.settings.enabled = True
        task_def.settings.stop_if_going_on_batteries = False

        start_time = datetime.now() + timedelta(minutes=5)
        trigger: TimeTrigger = task_def.triggers.create(TimeTrigger)
        trigger.start_boundary = start_time

        # Create action
        action: ExecAction = task_def.actions.create(ExecAction)
        action.id = "DO NOTHING"
        action.path = Path("cmd.exe")
        action.arguments = '/c "exit"'

        try:
            task_name: str = "Test Task"

            folder.register_task_definition(
                task_name,
                task_def,
                Creation.CREATE_OR_UPDATE,
                "",  # No user
                "",  # No password
                LogonType.NONE,
            )

            folder.get_task(task_name)
            folder.delete_task(task_name)
            with pytest.raises(TaskNotFoundError):
                folder.delete_task(task_name)
        finally:
            with contextlib.suppress(TaskFolderNotFoundError):
                ROOT.delete_folder(folder_name)

    def test_get_folder(self) -> None:
        """Test for TaskFolder.get_folder()."""
        name: str = "Folder"

        with contextlib.suppress(TaskFolderExistsError):
            folder: TaskFolder = ROOT.create_folder(name)
        try:
            found: TaskFolder = ROOT.get_folder(name)
            assert isinstance(folder, TaskFolder)
            assert folder == found
        finally:
            with contextlib.suppress(TaskFolderNotFoundError):
                ROOT.delete_folder(name)

    def test_get_folder_not_exists(self) -> None:
        """Test for TaskFolder.get_folder() with a folder that does not exist."""
        name: str = "FolderNotExists"

        with pytest.raises(TaskFolderNotFoundError):
            ROOT.get_folder(name)

    def test_get_folders(self) -> None:
        """Test for TaskFolder.get_folders()."""
        parent_name: str = "TestGetFolders"
        name: str = "Folder{i}"

        with contextlib.suppress(TaskFolderExistsError):
            parent_folder: TaskFolder = ROOT.create_folder(parent_name)
        try:
            for i in range(5):
                parent_folder.create_folder(name.format(i=i))

            found: TaskFolderCollection = parent_folder.get_folders()
            assert isinstance(found, TaskFolderCollection)
        finally:
            for i in range(5):
                with contextlib.suppress(TaskFolderNotFoundError):
                    parent_folder.delete_folder(name.format(i=i))
            with contextlib.suppress(TaskFolderNotFoundError):
                ROOT.delete_folder(parent_name)

    def test_get_security_descriptor(self) -> None:
        """Test for TaskFolder.get_security_descriptor()."""
        # TODO(Ryan): test_get_security_descriptor

    def test_get_task(self) -> None:
        """Test for TaskFolder.get_task()."""
        # TODO(Ryan): test_get_task

    def test_get_tasks(self) -> None:
        """Test for TaskFolder.get_tasks()."""
        # TODO(Ryan): test_get_tasks

    def test_register_task(self) -> None:
        """Test for TaskFolder.register_task()."""
        # TODO(Ryan): test_register_task

    def test_register_task_definition(self) -> None:
        """Test for TaskFolder.register_task_definition()."""
        # TODO(Ryan): test_register_task_definition

    def test_set_security_description(self) -> None:
        """Test for TaskFolder.set_security_description()."""
        # TODO(Ryan): test_set_security_description


class TestTaskFolderCollection:
    """Tests for TaskFolderCollection."""

    parent_name: str = "TestFolderCollection"
    name: str = "Folder{i}"
    count: int = 5

    @pytest.fixture(scope="class")
    def collection(self) -> Generator[TaskFolderCollection, Any, None]:
        """Collection fixture."""  # noqa: D401
        parent_folder: TaskFolder
        try:
            parent_folder = ROOT.create_folder(self.parent_name)
        except TaskFolderExistsError:
            parent_folder = ROOT.get_folder(self.parent_name)

        for i in range(self.count):
            i += 1
            with contextlib.suppress(TaskFolderExistsError):
                parent_folder.create_folder(self.name.format(i=i))

        yield parent_folder.get_folders()

        for i in range(self.count):
            i += 1
            with contextlib.suppress(TaskFolderNotFoundError):
                parent_folder.delete_folder(self.name.format(i=i))

        with contextlib.suppress(TaskFolderNotFoundError):
            ROOT.delete_folder(self.parent_name)

    def test_dunder_len(self, collection: TaskFolderCollection) -> None:
        """Test for TaskFolderCollection.__len__()."""
        assert len(collection) == self.count

    @pytest.mark.parametrize("index", [i + 1 for i in range(count)])
    def test_dunder_getitem(self, collection: TaskFolderCollection, index: int) -> None:
        """Test for TaskFolderCollection.__getitem__()."""
        folder: TaskFolder = collection[index]

        assert isinstance(folder, TaskFolder)
        assert folder.name == self.name.format(i=index)

    @pytest.mark.parametrize("index", [0, count + 1])
    def test_dunder_getitem_out_of_range(
        self, collection: TaskFolderCollection, index: int
    ) -> None:
        """Test for TaskFolderCollection.__getitem__()."""
        with pytest.raises(IndexError):
            collection.__getitem__(index)

    def test_dunder_iter(self, collection: TaskFolderCollection) -> None:
        """Test for TaskFolderCollection.__iter__()."""
        iterator: Iterator[TaskFolder] = iter(collection)
        assert isinstance(iterator, Iterator)

        member: TaskFolder
        for member in collection:
            assert isinstance(member, TaskFolder)

    @pytest.mark.parametrize("index", [i + 1 for i in range(count)])
    def test_dunder_contains(self, collection: TaskFolderCollection, index: int) -> None:
        """Test for TaskFolderCollection.__contains__()."""
        assert collection[index] in collection

    def test_count(self, collection: TaskFolderCollection) -> None:
        """Test for TaskFolderCollection.count."""
        assert collection.count == self.count

    @pytest.mark.parametrize("index", [i + 1 for i in range(count)])
    def test_item(self, collection: TaskFolderCollection, index: int) -> None:
        """Test for TaskFolderCollection.item()."""
        folder: TaskFolder = collection.item(index)

        assert isinstance(folder, TaskFolder)
        assert folder.name == self.name.format(i=index)

    @pytest.mark.parametrize("index", [0, count + 1])
    def test_item_out_of_range(self, collection: TaskFolderCollection, index: int) -> None:
        """Test for TaskFolderCollection.item()."""
        with pytest.raises(IndexError):
            collection.item(index)


class TestTaskDefinition:
    """Tests for TestTaskDefinition."""

    @pytest.fixture
    def definition(self) -> TaskDefinition:
        """Definition fixture."""
        return SERVICE.new_task()

    def test_actions(self, definition: TaskDefinition) -> None:
        """Test for TaskFolderCollection.actions."""
        obj: ActionCollection = definition.actions

        assert isinstance(obj, ActionCollection)

    def test_data(self, definition: TaskDefinition) -> None:
        """Test for TaskFolderCollection.data."""
        obj: str = definition.data

        assert isinstance(obj, str)

    def test_principal(self, definition: TaskDefinition) -> None:
        """Test for TaskFolderCollection.principal."""
        obj: Principal = definition.principal

        assert isinstance(obj, Principal)

    def test_registration_info(self, definition: TaskDefinition) -> None:
        """Test for TaskFolderCollection.registration_info."""
        obj: RegistrationInfo = definition.registration_info

        assert isinstance(obj, RegistrationInfo)

    def test_settings(self, definition: TaskDefinition) -> None:
        """Test for TaskFolderCollection.settings."""
        obj: TaskSettings = definition.settings

        assert isinstance(obj, TaskSettings)

    def test_triggers(self, definition: TaskDefinition) -> None:
        """Test for TaskFolderCollection.triggers."""
        obj: TriggerCollection = definition.triggers

        assert isinstance(obj, TriggerCollection)

    def test_xml_text(self, definition: TaskDefinition) -> None:
        """Test for TaskFolderCollection.xml_text."""
        obj: str = definition.xml_text

        assert isinstance(obj, str)


class TestRunningTask:
    """Tests for RunningTask."""

    # TODO(Ryan): TestRunningTask


class TestRunningTaskCollection:
    """Tests for RunningTaskCollection."""

    # TODO(Ryan): TestRunningTaskCollection


class TestRegisteredTask:
    """Tests for RegisteredTask."""

    task_name: str = "TestRegisteredTask"

    @pytest.fixture(scope="class")
    def definition(self) -> TaskDefinition:
        """Definition fixture."""
        definition: TaskDefinition = SERVICE.new_task()

        definition.registration_info.description = "Description"
        definition.principal.id = "Author"
        definition.settings.enabled = True
        definition.settings.stop_if_going_on_batteries = False

        trigger: TimeTrigger = definition.triggers.create(TimeTrigger)
        trigger.start_boundary = datetime.now() + timedelta(hours=1)

        action: ExecAction = definition.actions.create(ExecAction)
        action.id = "DO NOTHING"
        action.path = Path("cmd.exe")
        action.arguments = '/c "exit"'

        return definition

    @pytest.fixture(scope="class")
    def task(self, definition: TaskDefinition) -> Generator[RegisteredTask]:
        """Task fixture."""
        yield ROOT.register_task_definition(
            self.task_name,
            definition,
            Creation.CREATE_OR_UPDATE,
            "",
            "",
            LogonType.NONE,
        )

        with contextlib.suppress(TaskNotFoundError):
            ROOT.delete_task(self.task_name)

    def test_definition(self, task: RegisteredTask) -> None:
        """Test for RegisteredTask.definition."""
        obj: TaskDefinition = task.definition

        assert isinstance(obj, TaskDefinition)

    def test_enabled(self, definition: TaskDefinition, task: RegisteredTask) -> None:
        """Test for RegisteredTask.enabled."""
        obj: bool = task.enabled

        assert isinstance(obj, bool)
        assert obj == definition.settings.enabled

    def test_last_run_time(self, task: RegisteredTask) -> None:
        """Test for RegisteredTask.last_run_time."""
        obj: datetime = task.last_run_time

        assert isinstance(obj, datetime)

    def test_last_task_result(self, task: RegisteredTask) -> None:
        """Test for RegisteredTask.last_task_result."""
        obj: int = task.last_task_result

        assert isinstance(obj, int)

    def test_name(self, task: RegisteredTask) -> None:
        """Test for RegisteredTask.name."""
        obj: str = task.name

        assert isinstance(obj, str)
        assert obj == self.task_name

    def test_next_run_time(self, task: RegisteredTask) -> None:
        """Test for RegisteredTask.next_run_time."""
        obj: datetime = task.next_run_time

        assert isinstance(obj, datetime)

    def test_number_of_missed_runs(self, task: RegisteredTask) -> None:
        """Test for RegisteredTask.number_of_missed_runs."""
        obj: int = task.number_of_missed_runs

        assert isinstance(obj, int)

    def test_path(self, task: RegisteredTask) -> None:
        """Test for RegisteredTask.path."""
        obj: str = task.path

        assert isinstance(obj, str)

    def test_state(self, task: RegisteredTask) -> None:
        """Test for RegisteredTask.state."""
        obj: State = task.state

        assert isinstance(obj, State)

    def test_xml(self, task: RegisteredTask) -> None:
        """Test for RegisteredTask.xml."""
        obj: str = task.xml

        assert isinstance(obj, str)

    def test_get_instances(self, task: RegisteredTask) -> None:
        """Test for RegisteredTask.get_instances()."""
        obj: RunningTaskCollection = task.get_instances()

        assert isinstance(obj, RunningTaskCollection)

    def test_get_run_times(self, task: RegisteredTask) -> None:
        """Test for RegisteredTask.get_run_times()."""
        start: datetime = datetime.now()
        end: datetime = start + relativedelta(years=1)

        obj: RunningTaskCollection = task.get_run_times(start, end)

        assert isinstance(obj, RunningTaskCollection)

    def test_get_security_descriptor(self, task: RegisteredTask) -> None:
        """Test for RegisteredTask.get_security_descriptor()."""
        obj: str = task.get_security_descriptor(
            SecurityInformation.OWNER
            | SecurityInformation.GROUP
            | SecurityInformation.DACL
            | SecurityInformation.SACL
            | SecurityInformation.LABEL
        )

        assert isinstance(obj, str)

    def test_run(self, task: RegisteredTask) -> None:
        """Test for RegisteredTask.run()."""
        obj: RunningTask = task.run(None)

        assert isinstance(obj, RunningTask)

    def test_run_ex(self, task: RegisteredTask) -> None:
        """Test for RegisteredTask.run_ex()."""
        obj: RunningTask = task.run_ex(None, RunFlags.NO_FLAGS, 0)

        assert isinstance(obj, RunningTask)

    def test_set_security_descriptor(self, task: RegisteredTask) -> None:
        """Test for RegisteredTask.set_security_descriptor()."""
        security_descriptor: str = task.get_security_descriptor(
            SecurityInformation.OWNER
            | SecurityInformation.GROUP
            | SecurityInformation.DACL
            | SecurityInformation.SACL
            | SecurityInformation.LABEL
        )
        flags: Creation = Creation.CREATE_OR_UPDATE
        task.set_security_descriptor(security_descriptor, flags)

    def test_stop(self, task: RegisteredTask) -> None:
        """Test for RegisteredTask.stop()."""
        task.run(None)
        task.stop()


class TestRegisteredTaskCollection:
    """Tests for RegisteredTaskCollection."""

    parent_name: str = "TestRegisteredTaskCollection"
    name: str = "Task{i}"
    count: int = 5

    @pytest.fixture(scope="class")
    def collection(self) -> Generator[RegisteredTaskCollection, Any, None]:
        """Collection fixture."""  # noqa: D401
        parent_folder: TaskFolder
        try:
            parent_folder = ROOT.create_folder(self.parent_name)
        except TaskFolderExistsError:
            parent_folder = ROOT.get_folder(self.parent_name)

        definition: TaskDefinition = SERVICE.new_task()

        definition.registration_info.description = "Description"
        definition.settings.enabled = True
        definition.settings.stop_if_going_on_batteries = False

        trigger: TimeTrigger = definition.triggers.create(TimeTrigger)
        trigger.start_boundary = datetime.now() + timedelta(minutes=5)

        action: ExecAction = definition.actions.create(ExecAction)
        action.id = "DO NOTHING"
        action.path = Path("cmd.exe")
        action.arguments = '/c "exit"'

        for i in range(self.count):
            i += 1
            with contextlib.suppress(TaskFolderExistsError):
                parent_folder.register_task_definition(
                    self.name.format(i=i),
                    definition,
                    Creation.CREATE_OR_UPDATE,
                    "",
                    "",
                    LogonType.NONE,
                )

        yield parent_folder.get_tasks()

        for i in range(self.count):
            i += 1
            with contextlib.suppress(TaskFolderNotFoundError):
                parent_folder.delete_task(self.name.format(i=i))

        with contextlib.suppress(TaskFolderNotFoundError):
            ROOT.delete_folder(self.parent_name)

    def test_dunder_len(self, collection: RegisteredTaskCollection) -> None:
        """Test for RegisteredTaskCollection.__len__()."""
        assert len(collection) == self.count

    @pytest.mark.parametrize("index", [i + 1 for i in range(count)])
    def test_dunder_getitem(self, collection: RegisteredTaskCollection, index: int) -> None:
        """Test for RegisteredTaskCollection.__getitem__()."""
        task: RegisteredTask = collection[index]

        assert isinstance(task, RegisteredTask)
        assert task.name == self.name.format(i=index)

    @pytest.mark.parametrize("index", [0, count + 1])
    def test_dunder_getitem_out_of_range(
        self, collection: RegisteredTaskCollection, index: int
    ) -> None:
        """Test for RegisteredTaskCollection.__getitem__()."""
        with pytest.raises(IndexError):
            collection.__getitem__(index)

    def test_dunder_iter(self, collection: RegisteredTaskCollection) -> None:
        """Test for RegisteredTaskCollection.__iter__()."""
        iterator: Iterator[RegisteredTask] = iter(collection)
        assert isinstance(iterator, Iterator)

        for member in collection:
            assert isinstance(member, RegisteredTask)

    @pytest.mark.parametrize("index", [i + 1 for i in range(count)])
    def test_dunder_contains(self, collection: RegisteredTaskCollection, index: int) -> None:
        """Test for RegisteredTaskCollection.__contains__()."""
        assert collection[index] in collection

    def test_count(self, collection: RegisteredTaskCollection) -> None:
        """Test for RegisteredTaskCollection.count."""
        assert collection.count == self.count

    @pytest.mark.parametrize("index", [i + 1 for i in range(count)])
    def test_item(self, collection: RegisteredTaskCollection, index: int) -> None:
        """Test for RegisteredTaskCollection.item()."""
        task: RegisteredTask = collection.item(index)

        assert isinstance(task, RegisteredTask)
        assert task.name == self.name.format(i=index)

    @pytest.mark.parametrize("index", [0, count + 1])
    def test_item_out_of_range(self, collection: RegisteredTaskCollection, index: int) -> None:
        """Test for RegisteredTaskCollection.item()."""
        with pytest.raises(IndexError):
            collection.item(index)


class TestTaskVariables:
    """Tests for TaskVariables."""

    # TODO(Ryan): TestTaskVariables


# noinspection PyCompatibility
class TestRegistrationInfo:
    """Tests for RegistrationInfo."""

    @pytest.fixture
    def info(self) -> RegistrationInfo:
        """Info fixture."""
        task_def: TaskDefinition = SERVICE.new_task()
        return task_def.registration_info

    @pytest.mark.parametrize("expected", ["Author"])
    def test_author(self, info: RegistrationInfo, expected: str) -> None:
        """Test for RegistrationInfo.author."""
        info.author = expected

        obj: str = info.author

        assert isinstance(obj, str)
        assert obj == expected

    @pytest.mark.parametrize("expected", [None, datetime.now()])
    def test_date(self, info: RegistrationInfo, expected: datetime | None) -> None:
        """Test for RegistrationInfo.date."""
        info.date = expected

        obj: datetime | None = info.date

        assert isinstance(obj, datetime | None)
        assert obj == expected

    @pytest.mark.parametrize("expected", ["Description"])
    def test_description(self, info: RegistrationInfo, expected: str) -> None:
        """Test for RegistrationInfo.description."""
        info.description = expected

        obj: str = info.description

        assert isinstance(obj, str)
        assert obj == expected

    @pytest.mark.parametrize("expected", ["Documentation"])
    def test_documentation(self, info: RegistrationInfo, expected: str) -> None:
        """Test for RegistrationInfo.documentation."""
        info.documentation = expected

        obj: str = info.documentation

        assert isinstance(obj, str)
        assert obj == expected

    @pytest.mark.parametrize("expected", [None, "Security Descriptor"])
    def test_security_descriptor(self, info: RegistrationInfo, expected: str | None) -> None:
        """Test for RegistrationInfo.security_descriptor."""
        info.security_descriptor = expected

        obj: str | None = info.security_descriptor

        assert isinstance(obj, str | None)
        assert obj == expected

    @pytest.mark.parametrize("expected", ["Source"])
    def test_source(self, info: RegistrationInfo, expected: str) -> None:
        """Test for RegistrationInfo.source."""
        info.source = expected

        obj: str = info.source

        assert isinstance(obj, str)
        assert obj == expected

    @pytest.mark.parametrize("expected", ["URI"])
    def test_uri(self, info: RegistrationInfo, expected: str) -> None:
        """Test for RegistrationInfo.uri."""
        info.uri = expected

        obj: str = info.uri

        assert isinstance(obj, str)
        assert obj == expected

    @pytest.mark.parametrize("expected", ["Version"])
    def test_version(self, info: RegistrationInfo, expected: str) -> None:
        """Test for RegistrationInfo.version."""
        info.version = expected

        obj: str = info.version

        assert isinstance(obj, str)
        assert obj == expected

    def test_xml_text(self, info: RegistrationInfo) -> None:
        """Test for RegistrationInfo.xml_text."""
        obj: str = info.xml_text

        assert isinstance(obj, str)


class TestRepetitionPattern:
    """Tests for RepetitionPattern."""

    @pytest.fixture
    def pattern(self) -> RepetitionPattern:
        """Pattern fixture."""
        task_def: TaskDefinition = SERVICE.new_task()
        trigger: Trigger = task_def.triggers.create(EventTrigger)
        return trigger.repetition

    @pytest.mark.parametrize("expected", [relativedelta(seconds=30)])
    def test_duration(self, pattern: RepetitionPattern, expected: relativedelta) -> None:
        """Test for RepetitionPattern.duration."""
        pattern.duration = expected

        obj: relativedelta = pattern.duration

        assert isinstance(obj, relativedelta)
        assert obj == expected

    @pytest.mark.parametrize("expected", [relativedelta(seconds=30)])
    def test_interval(self, pattern: RepetitionPattern, expected: relativedelta) -> None:
        """Test for RepetitionPattern.interval."""
        pattern.interval = expected

        obj: relativedelta = pattern.interval

        assert isinstance(obj, relativedelta)
        assert obj == expected

    @pytest.mark.parametrize("expected", [True, False])
    def test_stop_at_duration_end(self, pattern: RepetitionPattern, expected: bool) -> None:
        """Test for RepetitionPattern.stop_at_duration_end."""
        pattern.stop_at_duration_end = expected

        obj: bool = pattern.stop_at_duration_end

        assert isinstance(obj, bool)
        assert obj == expected


class TestPrincipal:
    """Tests for Principal."""

    @pytest.fixture
    def principal(self) -> Principal:
        """Principal fixture."""
        task_def: TaskDefinition = SERVICE.new_task()
        return task_def.principal

    @pytest.mark.parametrize("expected", ["Display Name"])
    def test_display_name(self, principal: Principal, expected: str) -> None:
        """Test for Principal.display_name."""
        principal.display_name = expected

        obj: str = principal.display_name

        assert isinstance(obj, str)
        assert obj == expected

    @pytest.mark.parametrize("expected", ["Group ID"])
    def test_group_id(self, principal: Principal, expected: str) -> None:
        """Test for Principal.group_id."""
        principal.group_id = expected

        obj: str = principal.group_id

        assert isinstance(obj, str)
        assert obj == expected

    @pytest.mark.parametrize("expected", ["ID"])
    def test_id(self, principal: Principal, expected: str) -> None:
        """Test for Principal.id."""
        principal.id = expected

        obj: str = principal.id

        assert isinstance(obj, str)
        assert obj == expected

    @pytest.mark.parametrize("expected", list(LogonType)[1:])
    def test_logon_type(self, principal: Principal, expected: LogonType) -> None:
        """Test for Principal.logon_type."""
        principal.logon_type = expected

        obj: LogonType = principal.logon_type

        assert isinstance(obj, LogonType)
        assert obj == expected

    @pytest.mark.parametrize("expected", list(RunLevel))
    def test_run_level(self, principal: Principal, expected: RunLevel) -> None:
        """Test for Principal.run_level."""
        principal.run_level = expected

        obj: RunLevel = principal.run_level

        assert isinstance(obj, RunLevel)
        assert obj == expected

    @pytest.mark.parametrize("expected", ["User ID"])
    def test_user_id(self, principal: Principal, expected: str) -> None:
        """Test for Principal.user_id."""
        principal.user_id = expected

        obj: str = principal.user_id

        assert isinstance(obj, str)
        assert obj == expected


class TestTaskNamedValuePair:
    """Tests for TaskNamedValuePair."""

    @pytest.fixture
    def pair(self) -> TaskNamedValuePair:
        """Pair fixture."""
        task_def: TaskDefinition = SERVICE.new_task()
        trigger: EventTrigger = task_def.triggers.create(EventTrigger)
        return trigger.value_queries.create("Name", "Value")

    def test_dunder_len(self, pair: TaskNamedValuePair) -> None:
        """Test for TaskNamedValuePair.__len__()."""
        assert len(pair) == 2

    def test_dunder_getitem(self, pair: TaskNamedValuePair) -> None:
        """Test for TaskNamedValuePair.__getitem__()."""
        assert pair[0] == "Name"
        assert pair[1] == "Value"

    @pytest.mark.parametrize("index", [-1, 2])
    def test_dunder_getitem_out_of_range(self, pair: TaskNamedValuePair, index: int) -> None:
        """Test for TaskNamedValuePair.__getitem__()."""
        with pytest.raises(IndexError):
            pair.__getitem__(index)

    def test_dunder_iter(self, pair: TaskNamedValuePair) -> None:
        """Test for TaskNamedValuePair.__iter__()."""
        iterator: Iterator[str] = iter(pair)
        assert isinstance(iterator, Iterator)

        for obj in pair:
            assert isinstance(obj, str)

    @pytest.mark.parametrize("value", ["Name", "Value"])
    def test_dunder_contains(self, pair: TaskNamedValuePair, value: str) -> None:
        """Test for TaskNamedValuePair.__contains__()."""
        assert value in pair

    @pytest.mark.parametrize("expected", [f"Name{i}" for i in range(5)])
    def test_name(self, pair: TaskNamedValuePair, expected: str) -> None:
        """Test for TaskNamedValuePair.name."""
        pair.name = expected

        obj: str = pair.name

        assert isinstance(obj, str)
        assert obj == expected

    @pytest.mark.parametrize("expected", [f"Value{i}" for i in range(5)])
    def test_value(self, pair: TaskNamedValuePair, expected: str) -> None:
        """Test for TaskNamedValuePair.obj."""
        pair.obj = expected

        obj: str = pair.obj

        assert isinstance(obj, str)
        assert obj == expected


class TestTaskNamedValueCollection:
    """Tests for TaskNamedValueCollection."""

    name: str = "Name{i}"
    value: str = "Value{i}"
    count: int = 5

    @pytest.fixture
    def collection(self) -> TaskNamedValueCollection:
        """Collection fixture."""  # noqa: D401
        task_def: TaskDefinition = SERVICE.new_task()
        trigger: EventTrigger = task_def.triggers.create(EventTrigger)
        collection: TaskNamedValueCollection = trigger.value_queries

        for i in range(self.count):
            i += 1
            collection.create(self.name.format(i=i), self.value.format(i=i))

        return collection

    def test_dunder_len(self, collection: TaskNamedValueCollection) -> None:
        """Test for TaskNamedValueCollection.__len__()."""
        assert len(collection) == self.count

    @pytest.mark.parametrize("index", [i + 1 for i in range(count)])
    def test_dunder_getitem(self, collection: TaskNamedValueCollection, index: int) -> None:
        """Test for TaskNamedValueCollection.__getitem__()."""
        pair: TaskNamedValuePair = collection[index]

        assert isinstance(pair, TaskNamedValuePair)
        assert pair.name == self.name.format(i=index)
        assert pair.value == self.value.format(i=index)

    @pytest.mark.parametrize("index", [0, count + 1])
    def test_dunder_getitem_out_of_range(
        self, collection: TaskNamedValueCollection, index: int
    ) -> None:
        """Test for TaskNamedValueCollection.__getitem__()."""
        with pytest.raises(IndexError):
            collection.__getitem__(index)

    def test_dunder_iter(self, collection: TaskNamedValueCollection) -> None:
        """Test for TaskNamedValueCollection.__iter__()."""
        iterator: Iterator[TaskNamedValuePair] = iter(collection)
        assert isinstance(iterator, Iterator)

        for member in collection:
            assert isinstance(member, TaskNamedValuePair)

    @pytest.mark.parametrize("index", [i + 1 for i in range(count)])
    def test_dunder_contains(self, collection: TaskNamedValueCollection, index: int) -> None:
        """Test for TaskNamedValueCollection.__contains__()."""
        assert collection[index] in collection

    def test_count(self, collection: TaskNamedValueCollection) -> None:
        """Test for TaskNamedValueCollection.count."""
        assert collection.count == self.count

    @pytest.mark.parametrize("index", [i + 1 for i in range(count)])
    def test_item(self, collection: TaskNamedValueCollection, index: int) -> None:
        """Test for TaskNamedValueCollection.item()."""
        pair: TaskNamedValuePair = collection.item(index)

        assert isinstance(pair, TaskNamedValuePair)
        assert pair.name == self.name.format(i=index)
        assert pair.value == self.value.format(i=index)

    @pytest.mark.parametrize("index", [0, count + 1])
    def test_item_out_of_range(self, collection: TaskNamedValueCollection, index: int) -> None:
        """Test for TaskNamedValueCollection.item()."""
        with pytest.raises(IndexError):
            collection.item(index)

    def test_clear(self, collection: TaskNamedValueCollection) -> None:
        """Test for TaskNamedValueCollection.clear()."""
        assert collection.count == self.count

        collection.clear()

        assert collection.count == 0

    def test_create(self, collection: TaskNamedValueCollection) -> None:
        """Test for TaskNamedValueCollection.create()."""
        assert collection.count == self.count

        pair: TaskNamedValuePair = collection.create(self.name.format(i=6), self.value.format(i=6))

        assert isinstance(pair, TaskNamedValuePair)
        assert collection.count == self.count + 1

    def test_remove(self, collection: TaskNamedValueCollection) -> None:
        """Test for TaskNamedValueCollection.remove()."""
        assert collection.count == self.count

        collection.remove(self.count)

        assert collection.count == self.count - 1


class TestAction:
    """Tests for Action."""

    @pytest.fixture
    def action(self) -> Action:
        """Action fixture."""  # noqa: D401
        task_def: TaskDefinition = SERVICE.new_task()
        return task_def.actions.create(ExecAction)

    def test_id(self, action: Action) -> None:
        """Test for Action.id."""
        expected: str = "ID"
        action.id = expected

        obj: str = action.id
        assert isinstance(obj, str)
        assert obj == expected

    def test_type(self, action: Action) -> None:
        """Test for Action.type."""
        obj: int = action.type

        assert isinstance(obj, int)
        assert obj == ExecAction.type_value


class TestActionCollection:
    """Tests for ActionCollection."""

    values: Sequence[type[Action]] = [ExecAction, ComHandlerAction, EmailAction, ShowMessageAction]
    count: int = len(values)

    @pytest.fixture
    def collection(self) -> ActionCollection:
        """Collection fixture."""  # noqa: D401
        task_def: TaskDefinition = SERVICE.new_task()
        collection: ActionCollection = task_def.actions

        action: type[Action]
        for action in self.values:
            collection.create(action)

        return collection

    def test_dunder_len(self, collection: ActionCollection) -> None:
        """Test for ActionCollection.__len__()."""
        assert len(collection) == self.count

    @pytest.mark.parametrize("index", [i + 1 for i in range(count)])
    def test_dunder_getitem(self, collection: ActionCollection, index: int) -> None:
        """Test for ActionCollection.__getitem__()."""
        action: Action = collection[index]

        assert isinstance(action, self.values[index - 1])

    @pytest.mark.parametrize("index", [0, count + 1])
    def test_dunder_getitem_out_of_range(self, collection: ActionCollection, index: int) -> None:
        """Test for ActionCollection.__getitem__()."""
        with pytest.raises(IndexError):
            collection.__getitem__(index)

    def test_dunder_iter(self, collection: ActionCollection) -> None:
        """Test for ActionCollection.__iter__()."""
        iterator: Iterator[Action] = iter(collection)
        assert isinstance(iterator, Iterator)

        for member in collection:
            assert isinstance(member, Action)

    @pytest.mark.parametrize("index", [i + 1 for i in range(count)])
    def test_dunder_contains(self, collection: ActionCollection, index: int) -> None:
        """Test for ActionCollection.__contains__()."""
        assert collection[index] in collection

    @pytest.mark.parametrize("expected", ["Context"])
    def test_context(self, collection: ActionCollection, expected: str) -> None:
        """Test for ActionCollection.context."""
        collection.context = expected

        obj: str = collection.context

        assert isinstance(obj, str)
        assert obj == expected

    def test_count(self, collection: ActionCollection) -> None:
        """Test for ActionCollection.count."""
        assert collection.count == self.count

    @pytest.mark.parametrize("index", [i + 1 for i in range(count)])
    def test_item(self, collection: ActionCollection, index: int) -> None:
        """Test for ActionCollection.item()."""
        action: Action = collection.item(index)

        assert isinstance(action, self.values[index - 1])

    @pytest.mark.parametrize("index", [0, count + 1])
    def test_dunder_item_out_of_range(self, collection: ActionCollection, index: int) -> None:
        """Test for ActionCollection.item()."""
        with pytest.raises(IndexError):
            collection.item(index)

    def test_clear(self, collection: ActionCollection) -> None:
        """Test for ActionCollection.clear()."""
        assert collection.count == self.count

        collection.clear()

        assert collection.count == 0

    def test_create(self, collection: ActionCollection) -> None:
        """Test for ActionCollection.create()."""
        assert collection.count == self.count

        action: Action = collection.create(ExecAction)

        assert isinstance(action, ExecAction)
        assert collection.count == self.count + 1

    def test_remove(self, collection: ActionCollection) -> None:
        """Test for ActionCollection.remove()."""
        assert collection.count == self.count

        collection.remove(self.count)

        assert collection.count == self.count - 1


class TestExecAction:
    """Tests for ExecAction."""

    @pytest.fixture
    def action(self) -> ExecAction:
        """Action fixture."""  # noqa: D401
        task_def: TaskDefinition = SERVICE.new_task()
        return task_def.actions.create(ExecAction)

    @pytest.mark.parametrize("expected", ["Arguments"])
    def test_arguments(self, action: ExecAction, expected: str) -> None:
        """Test for Action.arguments."""
        action.arguments = expected

        obj: str = action.arguments

        assert isinstance(obj, str)
        assert obj == expected

    @pytest.mark.parametrize("expected", [Path("path/to/executable.exe").resolve()])
    def test_path(self, action: ExecAction, expected: Path) -> None:
        """Test for Action.path."""
        action.path = expected

        obj: Path = action.path

        assert isinstance(obj, Path)
        assert obj == expected

    @pytest.mark.parametrize("expected", [Path("working/directory/path").resolve()])
    def test_working_directory(self, action: ExecAction, expected: Path) -> None:
        """Test for Action.working_directory."""
        action.working_directory = expected

        obj: Path = action.working_directory

        assert isinstance(obj, Path)
        assert obj == expected


class TestComHandlerAction:
    """Tests for ComHandlerAction."""

    @pytest.fixture
    def action(self) -> ComHandlerAction:
        """Action fixture."""  # noqa: D401
        task_def: TaskDefinition = SERVICE.new_task()
        return task_def.actions.create(ComHandlerAction)

    @pytest.mark.parametrize("expected", ["Class ID"])
    def test_class_id(self, action: ComHandlerAction, expected: str) -> None:
        """Test for ComHandlerAction.class_id."""
        action.class_id = expected

        obj: str = action.class_id

        assert isinstance(obj, str)
        assert obj == expected

    @pytest.mark.parametrize("expected", ["Data"])
    def test_data(self, action: ComHandlerAction, expected: str) -> None:
        """Test for ComHandlerAction.data."""
        action.data = expected

        obj: str = action.data

        assert isinstance(obj, str)
        assert obj == expected


class TestEmailAction:
    """Tests for EmailAction."""

    @pytest.fixture
    def action(self) -> EmailAction:
        """Action fixture."""  # noqa: D401
        task_def: TaskDefinition = SERVICE.new_task()
        return task_def.actions.create(EmailAction)

    @pytest.mark.parametrize("expected", [(), tuple(f"Attachment{i}" for i in range(5))])
    def test_attachments(self, action: EmailAction, expected: tuple[str, ...] | None) -> None:
        """Test for EmailAction.attachments."""
        action.attachments = expected

        obj: tuple[str, ...] = action.attachments

        assert isinstance(obj, tuple)
        assert obj == expected

    @pytest.mark.parametrize("expected", ["BCC"])
    def test_bcc(self, action: EmailAction, expected: str) -> None:
        """Test for EmailAction.bcc."""
        action.bcc = expected

        obj: str = action.bcc

        assert isinstance(obj, str)
        assert obj == expected

    @pytest.mark.parametrize("expected", ["Body"])
    def test_body(self, action: EmailAction, expected: str) -> None:
        """Test for EmailAction.body."""
        action.body = expected

        obj: str = action.body

        assert isinstance(obj, str)
        assert obj == expected

    @pytest.mark.parametrize("expected", ["CC"])
    def test_cc(self, action: EmailAction, expected: str) -> None:
        """Test for EmailAction.cc."""
        action.cc = expected

        obj: str = action.cc

        assert isinstance(obj, str)
        assert obj == expected

    @pytest.mark.parametrize("expected", ["From"])
    def test_from_(self, action: EmailAction, expected: str) -> None:
        """Test for EmailAction.from_."""
        action.from_ = expected

        obj: str = action.from_

        assert isinstance(obj, str)
        assert obj == expected

    def test_header_fields(self, action: EmailAction) -> None:
        """Test for EmailAction.header_fields."""
        obj: TaskNamedValueCollection = action.header_fields

        assert isinstance(obj, TaskNamedValueCollection)
        assert obj is action.header_fields

    @pytest.mark.parametrize("expected", ["Reply To"])
    def test_reply_to(self, action: EmailAction, expected: str) -> None:
        """Test for EmailAction.reply_to."""
        action.reply_to = expected

        obj: str = action.reply_to

        assert isinstance(obj, str)
        assert obj == expected

    @pytest.mark.parametrize("expected", ["Server"])
    def test_server(self, action: EmailAction, expected: str) -> None:
        """Test for EmailAction.server."""
        action.server = expected

        obj: str = action.server

        assert isinstance(obj, str)
        assert obj == expected

    @pytest.mark.parametrize("expected", ["Subject"])
    def test_subject(self, action: EmailAction, expected: str) -> None:
        """Test for EmailAction.subject."""
        action.subject = expected

        obj: str = action.subject

        assert isinstance(obj, str)
        assert obj == expected

    @pytest.mark.parametrize("expected", ["To"])
    def test_to(self, action: EmailAction, expected: str) -> None:
        """Test for EmailAction.to."""
        action.to = expected

        obj: str = action.to

        assert isinstance(obj, str)
        assert obj == expected


class TestShowMessageAction:
    """Tests for ShowMessageAction."""

    @pytest.fixture
    def action(self) -> ShowMessageAction:
        """Action fixture."""  # noqa: D401
        task_def: TaskDefinition = SERVICE.new_task()
        return task_def.actions.create(ShowMessageAction)

    @pytest.mark.parametrize("expected", ["Message Body"])
    def test_message_body(self, action: ShowMessageAction, expected: str) -> None:
        """Test for ShowMessageAction.message_body."""
        action.message_body = expected

        obj: str = action.message_body

        assert isinstance(obj, str)
        assert obj == expected

    @pytest.mark.parametrize("expected", ["Title"])
    def test_title(self, action: ShowMessageAction, expected: str) -> None:
        """Test for ShowMessageAction.title."""
        action.title = expected

        obj: str = action.title

        assert isinstance(obj, str)
        assert obj == expected


# noinspection PyCompatibility
class TestTrigger:
    """Tests for Trigger."""

    @pytest.fixture
    def trigger(self) -> Trigger:
        """Trigger fixture."""
        task_def: TaskDefinition = SERVICE.new_task()
        return task_def.triggers.create(EventTrigger)

    @pytest.mark.parametrize("expected", [False, True])
    def test_enabled(self, trigger: Trigger, expected: bool) -> None:
        """Test for Trigger.enabled."""
        trigger.enabled = expected

        obj: bool = trigger.enabled

        assert isinstance(obj, bool)
        assert obj == expected

    @pytest.mark.parametrize("expected", [None, datetime.now()])
    def test_end_boundary(self, trigger: Trigger, expected: datetime | None) -> None:
        """Test for Trigger.end_boundary."""
        trigger.end_boundary = expected

        obj: datetime | None = trigger.end_boundary

        assert isinstance(obj, datetime | None)
        assert obj == expected

    @pytest.mark.parametrize("expected", [relativedelta(days=3)])
    def test_execution_time_limit(self, trigger: Trigger, expected: relativedelta) -> None:
        """Test for Trigger.execution_time_limit."""
        trigger.execution_time_limit = expected

        obj: relativedelta = trigger.execution_time_limit

        assert isinstance(obj, relativedelta)
        assert obj == expected

    @pytest.mark.parametrize("expected", ["ID"])
    def test_id(self, trigger: Trigger, expected: str) -> None:
        """Test for Trigger.id."""
        trigger.id = expected

        obj: str = trigger.id

        assert isinstance(obj, str)
        assert obj == expected

    def test_repetition(self, trigger: Trigger) -> None:
        """Test for Trigger.repetition."""
        obj: RepetitionPattern = trigger.repetition

        assert isinstance(obj, RepetitionPattern)
        assert obj is trigger.repetition

    @pytest.mark.parametrize("expected", [None, datetime.now()])
    def test_start_boundary(self, trigger: Trigger, expected: datetime | None) -> None:
        """Test for Trigger.start_boundary."""
        trigger.start_boundary = expected

        obj: datetime | None = trigger.start_boundary

        assert isinstance(obj, datetime | None)
        assert obj == expected

    def test_type(self, trigger: Trigger) -> None:
        """Test for Trigger.type."""
        obj: int = trigger.type

        assert isinstance(obj, int)
        assert obj == EventTrigger.type_value


class TestTriggerCollection:
    """Tests for TriggerCollection."""

    values: Sequence[type[Trigger]] = [
        EventTrigger,
        TimeTrigger,
        DailyTrigger,
        WeeklyTrigger,
        MonthlyTrigger,
        MonthlyDOWTrigger,
        IdleTrigger,
        RegistrationTrigger,
        BootTrigger,
        LogonTrigger,
        SessionStateChangeTrigger,
    ]
    count: int = len(values)

    @pytest.fixture
    def collection(self) -> TriggerCollection:
        """Collection fixture."""  # noqa: D401
        task_def: TaskDefinition = SERVICE.new_task()
        collection: TriggerCollection = task_def.triggers

        trigger: type[Trigger]
        for trigger in self.values:
            collection.create(trigger)

        return collection

    def test_dunder_len(self, collection: TriggerCollection) -> None:
        """Test for TriggerCollection.__len__()."""
        assert len(collection) == self.count

    @pytest.mark.parametrize("index", [i + 1 for i in range(count)])
    def test_dunder_getitem(self, collection: TriggerCollection, index: int) -> None:
        """Test for TriggerCollection.__getitem__()."""
        trigger: Trigger = collection[index]

        assert isinstance(trigger, self.values[index - 1])

    @pytest.mark.parametrize("index", [0, count + 1])
    def test_dunder_getitem_out_of_range(self, collection: TriggerCollection, index: int) -> None:
        """Test for TriggerCollection.__getitem__()."""
        with pytest.raises(IndexError):
            collection.__getitem__(index)

    def test_dunder_iter(self, collection: TriggerCollection) -> None:
        """Test for TriggerCollection.__iter__()."""
        iterator: Iterator[Trigger] = iter(collection)
        assert isinstance(iterator, Iterator)

        for member in collection:
            assert isinstance(member, Trigger)

    @pytest.mark.parametrize("index", [i + 1 for i in range(count)])
    def test_dunder_contains(self, collection: TriggerCollection, index: int) -> None:
        """Test for TriggerCollection.__contains__()."""
        assert collection[index] in collection

    def test_count(self, collection: TriggerCollection) -> None:
        """Test for TriggerCollection.count."""
        assert collection.count == self.count

    @pytest.mark.parametrize("index", [i + 1 for i in range(count)])
    def test_item(self, collection: TriggerCollection, index: int) -> None:
        """Test for TriggerCollection.item()."""
        trigger: Trigger = collection.item(index)

        assert isinstance(trigger, self.values[index - 1])

    @pytest.mark.parametrize("index", [0, count + 1])
    def test_item_out_of_range(self, collection: TriggerCollection, index: int) -> None:
        """Test for TriggerCollection.item()."""
        with pytest.raises(IndexError):
            collection.item(index)

    def test_clear(self, collection: TriggerCollection) -> None:
        """Test for TriggerCollection.clear()."""
        assert collection.count == self.count

        collection.clear()

        assert collection.count == 0

    def test_create(self, collection: TriggerCollection) -> None:
        """Test for TriggerCollection.create()."""
        assert collection.count == self.count

        trigger: Trigger = collection.create(EventTrigger)

        assert isinstance(trigger, EventTrigger)
        assert collection.count == self.count + 1

    def test_remove(self, collection: TriggerCollection) -> None:
        """Test for TriggerCollection.remove()."""
        assert collection.count == self.count

        collection.remove(self.count)

        assert collection.count == self.count - 1


class TestEventTrigger:
    """Tests for EventTrigger."""

    @pytest.fixture
    def trigger(self) -> EventTrigger:
        """Trigger fixture."""
        task_def: TaskDefinition = SERVICE.new_task()
        return task_def.triggers.create(EventTrigger)

    @pytest.mark.parametrize("expected", [relativedelta(minutes=30)])
    def test_delay(self, trigger: EventTrigger, expected: relativedelta) -> None:
        """Test for EventTrigger.delay."""
        trigger.delay = expected

        obj: relativedelta = trigger.delay

        assert isinstance(obj, relativedelta)
        assert obj == expected

    @pytest.mark.parametrize("expected", ["Subscription"])
    def test_subscription(self, trigger: EventTrigger, expected: str) -> None:
        """Test for EventTrigger.subscription."""
        trigger.subscription = expected

        obj: str = trigger.subscription

        assert isinstance(obj, str)
        assert obj == expected

    def test_value_queries(self, trigger: EventTrigger) -> None:
        """Test for EventTrigger.value_queries."""
        obj: TaskNamedValueCollection = trigger.value_queries

        assert isinstance(obj, TaskNamedValueCollection)
        assert obj is trigger.value_queries


class TestTimeTrigger:
    """Tests for TimeTrigger."""

    @pytest.fixture
    def trigger(self) -> TimeTrigger:
        """Trigger fixture."""
        task_def: TaskDefinition = SERVICE.new_task()
        return task_def.triggers.create(TimeTrigger)

    @pytest.mark.parametrize("expected", [relativedelta(minutes=30)])
    def test_random_delay(self, trigger: TimeTrigger, expected: relativedelta) -> None:
        """Test for TimeTrigger.random_delay."""
        trigger.random_delay = expected

        obj: relativedelta = trigger.random_delay

        assert isinstance(obj, relativedelta)
        assert obj == expected


class TestDailyTrigger:
    """Tests for DailyTrigger."""

    @pytest.fixture
    def trigger(self) -> DailyTrigger:
        """Trigger fixture."""
        task_def: TaskDefinition = SERVICE.new_task()
        return task_def.triggers.create(DailyTrigger)

    @pytest.mark.parametrize("expected", list(range(1, 11)))
    def test_days_interval(self, trigger: DailyTrigger, expected: int) -> None:
        """Test for TimeTrigger.days_interval."""
        trigger.days_interval = expected

        obj: int = trigger.days_interval

        assert isinstance(obj, int)
        assert obj == expected


class TestWeeklyTrigger:
    """Tests for WeeklyTrigger."""

    @pytest.fixture
    def trigger(self) -> WeeklyTrigger:
        """Trigger fixture."""
        task_def: TaskDefinition = SERVICE.new_task()
        return task_def.triggers.create(WeeklyTrigger)

    @pytest.mark.parametrize("expected", list(DaysOfWeek))
    def test_days_of_week(self, trigger: WeeklyTrigger, expected: DaysOfWeek) -> None:
        """Test for WeeklyTrigger.days_of_week."""
        trigger.days_of_week = expected

        obj: DaysOfWeek = trigger.days_of_week

        assert isinstance(obj, DaysOfWeek)
        assert obj == expected

    @pytest.mark.parametrize("expected", list(range(1, 11)))
    def test_weeks_interval(self, trigger: WeeklyTrigger, expected: int) -> None:
        """Test for WeeklyTrigger.weeks_interval."""
        trigger.weeks_interval = expected

        obj: int = trigger.weeks_interval

        assert isinstance(obj, int)
        assert obj == expected


class TestMonthlyTrigger:
    """Tests for MonthlyTrigger."""

    @pytest.fixture
    def trigger(self) -> MonthlyTrigger:
        """Trigger fixture."""
        task_def: TaskDefinition = SERVICE.new_task()
        return task_def.triggers.create(MonthlyTrigger)

    @pytest.mark.parametrize("expected", list(DaysOfMonth)[:-1])
    def test_days_of_month(self, trigger: MonthlyTrigger, expected: DaysOfMonth) -> None:
        """Test for MonthlyTrigger.days_of_month."""
        trigger.days_of_month = expected

        obj: DaysOfMonth = trigger.days_of_month

        assert isinstance(obj, DaysOfMonth)
        assert obj == expected

    @pytest.mark.parametrize("expected", list(MonthsOfYear))
    def test_months_of_year(self, trigger: MonthlyTrigger, expected: MonthsOfYear) -> None:
        """Test for MonthlyTrigger.months_of_year."""
        trigger.months_of_year = expected

        obj: MonthsOfYear = trigger.months_of_year

        assert isinstance(obj, MonthsOfYear)
        assert obj == expected

    @pytest.mark.parametrize("expected", [False, True])
    def test_run_on_last_day_of_month(self, trigger: MonthlyTrigger, expected: bool) -> None:
        """Test for MonthlyTrigger.run_on_last_day_of_month."""
        trigger.run_on_last_day_of_month = expected

        obj: bool = trigger.run_on_last_day_of_month

        assert isinstance(obj, bool)
        assert obj == expected


class TestMonthlyDOWTrigger:
    """Tests for MonthlyDOWTrigger."""

    @pytest.fixture
    def trigger(self) -> MonthlyDOWTrigger:
        """Trigger fixture."""
        task_def: TaskDefinition = SERVICE.new_task()
        return task_def.triggers.create(MonthlyDOWTrigger)

    @pytest.mark.parametrize("expected", list(DaysOfWeek))
    def test_days_of_week(self, trigger: MonthlyDOWTrigger, expected: DaysOfWeek) -> None:
        """Test for MonthlyDOWTrigger.days_of_week."""
        trigger.days_of_week = expected

        obj: DaysOfWeek = trigger.days_of_week

        assert isinstance(obj, DaysOfWeek)
        assert obj == expected

    @pytest.mark.parametrize("expected", list(MonthsOfYear))
    def test_months_of_year(self, trigger: MonthlyDOWTrigger, expected: MonthsOfYear) -> None:
        """Test for MonthlyDOWTrigger.months_of_year."""
        trigger.months_of_year = expected

        obj: MonthsOfYear = trigger.months_of_year

        assert isinstance(obj, MonthsOfYear)
        assert obj == expected

    @pytest.mark.parametrize("expected", [False, True])
    def test_run_on_last_week_of_month(self, trigger: MonthlyDOWTrigger, expected: bool) -> None:
        """Test for MonthlyDOWTrigger.run_on_last_week_of_month."""
        trigger.run_on_last_week_of_month = expected

        obj: bool = trigger.run_on_last_week_of_month

        assert isinstance(obj, bool)
        assert obj == expected

    @pytest.mark.parametrize("expected", list(WeeksOfMonth))
    def test_weeks_of_month(self, trigger: MonthlyDOWTrigger, expected: WeeksOfMonth) -> None:
        """Test for MonthlyDOWTrigger.weeks_of_month."""
        trigger.weeks_of_month = expected

        obj: WeeksOfMonth = trigger.weeks_of_month

        assert isinstance(obj, WeeksOfMonth)
        assert obj == expected


class TestIdleTrigger:
    """Tests for IdleTrigger."""

    @pytest.fixture
    def trigger(self) -> IdleTrigger:
        """Trigger fixture."""
        task_def: TaskDefinition = SERVICE.new_task()
        return task_def.triggers.create(IdleTrigger)


class TestRegistrationTrigger:
    """Tests for RegistrationTrigger."""

    @pytest.fixture
    def trigger(self) -> RegistrationTrigger:
        """Trigger fixture."""
        task_def: TaskDefinition = SERVICE.new_task()
        return task_def.triggers.create(RegistrationTrigger)

    @pytest.mark.parametrize("expected", [relativedelta(hours=5)])
    def test_delay(self, trigger: RegistrationTrigger, expected: relativedelta) -> None:
        """Test for RegistrationTrigger.delay."""
        trigger.delay = expected

        obj: relativedelta = trigger.delay

        assert isinstance(obj, relativedelta)
        assert obj == expected


class TestBootTrigger:
    """Tests for BootTrigger."""

    @pytest.fixture
    def trigger(self) -> BootTrigger:
        """Trigger fixture."""
        task_def: TaskDefinition = SERVICE.new_task()
        return task_def.triggers.create(BootTrigger)

    @pytest.mark.parametrize("expected", [relativedelta(hours=5)])
    def test_delay(self, trigger: BootTrigger, expected: relativedelta) -> None:
        """Test for BootTrigger.delay."""
        trigger.delay = expected

        obj: relativedelta = trigger.delay

        assert isinstance(obj, relativedelta)
        assert obj == expected


class TestLogonTrigger:
    """Tests for LogonTrigger."""

    @pytest.fixture
    def trigger(self) -> LogonTrigger:
        """Trigger fixture."""
        task_def: TaskDefinition = SERVICE.new_task()
        return task_def.triggers.create(LogonTrigger)

    @pytest.mark.parametrize("expected", [relativedelta(hours=5)])
    def test_delay(self, trigger: LogonTrigger, expected: relativedelta) -> None:
        """Test for BootTrigger.delay."""
        trigger.delay = expected

        obj: relativedelta = trigger.delay

        assert isinstance(obj, relativedelta)
        assert obj == expected

    @pytest.mark.parametrize("expected", ["User ID"])
    def test_user_id(self, trigger: LogonTrigger, expected: str) -> None:
        """Test for LogonTrigger.user_id."""
        trigger.user_id = expected

        obj: str = trigger.user_id

        assert isinstance(obj, str)
        assert obj == expected


class TestSessionStateChangeTrigger:
    """Tests for SessionStateChangeTrigger."""

    @pytest.fixture
    def trigger(self) -> SessionStateChangeTrigger:
        """Trigger fixture."""
        task_def: TaskDefinition = SERVICE.new_task()
        return task_def.triggers.create(SessionStateChangeTrigger)

    @pytest.mark.parametrize("expected", [relativedelta(hours=5)])
    def test_delay(self, trigger: SessionStateChangeTrigger, expected: relativedelta) -> None:
        """Test for SessionStateChangeTrigger.delay."""
        trigger.delay = expected

        obj: relativedelta = trigger.delay

        assert isinstance(obj, relativedelta)
        assert obj == expected

    @pytest.mark.parametrize("expected", list(SessionStateChangeType))
    def test_state_change(
        self, trigger: SessionStateChangeTrigger, expected: SessionStateChangeType
    ) -> None:
        """Test for SessionStateChangeTrigger.state_change."""
        trigger.state_change = expected

        obj: SessionStateChangeType = trigger.state_change

        assert isinstance(obj, SessionStateChangeType)
        assert obj == expected

    @pytest.mark.parametrize("expected", ["User ID"])
    def test_user_id(self, trigger: SessionStateChangeTrigger, expected: str) -> None:
        """Test for SessionStateChangeTrigger.user_id."""
        trigger.user_id = expected

        obj: str = trigger.user_id

        assert isinstance(obj, str)
        assert obj == expected


# noinspection PyCompatibility
class TestTaskSettings:
    """Tests for TaskSettings."""

    @pytest.fixture
    def settings(self) -> TaskSettings:
        """Settings fixture."""  # noqa: D401
        task_def: TaskDefinition = SERVICE.new_task()
        return task_def.settings

    @pytest.mark.parametrize("expected", [False, True])
    def test_allow_demand_start(self, settings: TaskSettings, expected: bool) -> None:
        """Test for TaskSettings.allow_demand_start."""
        settings.allow_demand_start = expected

        obj: bool = settings.allow_demand_start

        assert isinstance(obj, bool)
        assert obj == expected

    @pytest.mark.parametrize("expected", [False, True])
    def test_allow_hard_terminate(self, settings: TaskSettings, expected: bool) -> None:
        """Test for TaskSettings.allow_hard_terminate."""
        settings.allow_hard_terminate = expected

        obj: bool = settings.allow_hard_terminate

        assert isinstance(obj, bool)
        assert obj == expected

    @pytest.mark.parametrize("expected", list(Compatibility))
    def test_compatibility(self, settings: TaskSettings, expected: Compatibility) -> None:
        """Test for TaskSettings.compatibility."""
        settings.compatibility = expected

        obj: Compatibility = settings.compatibility

        assert isinstance(obj, Compatibility)
        assert obj == expected

    @pytest.mark.parametrize("expected", [None, datetime.now()])
    def test_delete_expired_task_after(
        self, settings: TaskSettings, expected: datetime | None
    ) -> None:
        """Test for TaskSettings.delete_expired_task_after."""
        settings.delete_expired_task_after = expected

        obj: datetime | None = settings.delete_expired_task_after

        assert isinstance(obj, datetime | None)
        assert obj == expected

    @pytest.mark.parametrize("expected", [False, True])
    def test_disallow_start_if_on_batteries(self, settings: TaskSettings, expected: bool) -> None:
        """Test for TaskSettings.disallow_start_if_on_batteries."""
        settings.disallow_start_if_on_batteries = expected

        obj: bool = settings.disallow_start_if_on_batteries

        assert isinstance(obj, bool)
        assert obj == expected

    @pytest.mark.parametrize("expected", [False, True])
    def test_enabled(self, settings: TaskSettings, expected: bool) -> None:
        """Test for TaskSettings.enabled."""
        settings.enabled = expected

        obj: bool = settings.enabled

        assert isinstance(obj, bool)
        assert obj == expected

    @pytest.mark.parametrize("expected", [None, relativedelta(days=10)])
    def test_execution_time_limit(
        self, settings: TaskSettings, expected: relativedelta | None
    ) -> None:
        """Test for TaskSettings.execution_time_limit."""
        settings.execution_time_limit = expected

        obj: relativedelta | None = settings.execution_time_limit

        assert isinstance(obj, relativedelta | None)
        assert obj == expected

    @pytest.mark.parametrize("expected", [False, True])
    def test_hidden(self, settings: TaskSettings, expected: bool) -> None:
        """Test for TaskSettings.hidden."""
        settings.hidden = expected

        obj: bool = settings.hidden

        assert isinstance(obj, bool)
        assert obj == expected

    def test_idle_settings(self, settings: TaskSettings) -> None:
        """Test for TaskSettings.idle_settings."""
        obj: IdleSettings = settings.idle_settings

        assert isinstance(obj, IdleSettings)
        assert obj is settings.idle_settings

    @pytest.mark.parametrize("expected", list(InstancesPolicy))
    def test_multiple_instances(self, settings: TaskSettings, expected: InstancesPolicy) -> None:
        """Test for TaskSettings.multiple_instances."""
        settings.multiple_instances = expected

        obj: InstancesPolicy = settings.multiple_instances

        assert isinstance(obj, InstancesPolicy)
        assert obj == expected

    def test_network_settings(self, settings: TaskSettings) -> None:
        """Test for TaskSettings.network_settings."""
        obj: NetworkSettings = settings.network_settings

        assert isinstance(obj, NetworkSettings)
        assert obj is settings.network_settings

    @pytest.mark.parametrize("expected", list(range(11)))
    def test_priority(self, settings: TaskSettings, expected: int) -> None:
        """Test for TaskSettings.priority."""
        settings.priority = expected

        obj: int = settings.priority

        assert isinstance(obj, int)
        assert obj == expected

    @pytest.mark.parametrize("expected", [-1, 11])
    def test_priority_out_of_range(self, settings: TaskSettings, expected: int) -> None:
        """Test for TaskSettings.priority."""
        with pytest.raises(ValueError, match="Invalid Priority."):
            settings.priority = expected

    @pytest.mark.parametrize("expected", list(range(-1, 11)))
    def test_restart_count(self, settings: TaskSettings, expected: int) -> None:
        """Test for TaskSettings.restart_count."""
        settings.restart_count = expected

        obj: int = settings.restart_count

        assert isinstance(obj, int)
        assert obj == expected

    @pytest.mark.parametrize("expected", [None, relativedelta(hours=1)])
    def test_restart_interval(self, settings: TaskSettings, expected: relativedelta | None) -> None:
        """Test for TaskSettings.restart_interval."""
        settings.restart_interval = expected

        obj: relativedelta | None = settings.restart_interval

        assert isinstance(obj, relativedelta | None)
        assert obj == expected

    @pytest.mark.parametrize("expected", [False, True])
    def test_run_only_if_idle(self, settings: TaskSettings, expected: bool) -> None:
        """Test for TaskSettings.run_only_if_idle."""
        settings.run_only_if_idle = expected

        obj: bool = settings.run_only_if_idle

        assert isinstance(obj, bool)
        assert obj == expected

    @pytest.mark.parametrize("expected", [False, True])
    def test_run_only_if_network_available(self, settings: TaskSettings, expected: bool) -> None:
        """Test for TaskSettings.run_only_if_network_available."""
        settings.run_only_if_network_available = expected

        obj: bool = settings.run_only_if_network_available

        assert isinstance(obj, bool)
        assert obj == expected

    @pytest.mark.parametrize("expected", [False, True])
    def test_start_when_available(self, settings: TaskSettings, expected: bool) -> None:
        """Test for TaskSettings.start_when_available."""
        settings.start_when_available = expected

        obj: bool = settings.start_when_available

        assert isinstance(obj, bool)
        assert obj == expected

    @pytest.mark.parametrize("expected", [False, True])
    def test_stop_if_going_on_batteries(self, settings: TaskSettings, expected: bool) -> None:
        """Test for TaskSettings.stop_if_going_on_batteries."""
        settings.stop_if_going_on_batteries = expected

        obj: bool = settings.stop_if_going_on_batteries

        assert isinstance(obj, bool)
        assert obj == expected

    @pytest.mark.parametrize("expected", [False, True])
    def test_wake_to_run(self, settings: TaskSettings, expected: bool) -> None:
        """Test for TaskSettings.wake_to_run."""
        settings.wake_to_run = expected

        obj: bool = settings.wake_to_run

        assert isinstance(obj, bool)
        assert obj == expected

    def test_xml_text(self, settings: TaskSettings) -> None:
        """Test for TaskSettings.xml_text."""
        obj: str = settings.xml_text

        assert isinstance(obj, str)


class TestIdleSettings:
    """Tests for IdleSettings."""

    @pytest.fixture
    def settings(self) -> IdleSettings:
        """Settings fixture."""  # noqa: D401
        task_def: TaskDefinition = SERVICE.new_task()
        task_settings: TaskSettings = task_def.settings
        return task_settings.idle_settings

    @pytest.mark.parametrize("expected", [False, True])
    def test_restart_on_idle(self, settings: IdleSettings, expected: bool) -> None:
        """Test for IdleSettings.restart_on_idle."""
        settings.restart_on_idle = expected

        obj: bool = settings.restart_on_idle

        assert isinstance(obj, bool)
        assert obj == expected

    @pytest.mark.parametrize("expected", [False, True])
    def test_stop_on_idle_end(self, settings: IdleSettings, expected: bool) -> None:
        """Test for IdleSettings.stop_on_idle_end."""
        settings.stop_on_idle_end = expected

        obj: bool = settings.stop_on_idle_end

        assert isinstance(obj, bool)
        assert obj == expected


class TestNetworkSettings:
    """Tests for NetworkSettings."""

    @pytest.fixture
    def settings(self) -> NetworkSettings:
        """Settings fixture."""  # noqa: D401
        task_def: TaskDefinition = SERVICE.new_task()
        task_settings: TaskSettings = task_def.settings
        return task_settings.network_settings

    @pytest.mark.parametrize("expected", ["ID"])
    def test_id(self, settings: NetworkSettings, expected: str) -> None:
        """Test for NetworkSettings.id."""
        settings.id = expected

        obj: str = settings.id

        assert isinstance(obj, str)
        assert obj == expected

    @pytest.mark.parametrize("expected", ["Name"])
    def test_name(self, settings: NetworkSettings, expected: str) -> None:
        """Test for NetworkSettings.name."""
        settings.name = expected

        obj: str = settings.name

        assert isinstance(obj, str)
        assert obj == expected


# noinspection PyCompatibility
class TestFromDurationStr:
    """Tests for from_duration_str()."""

    value_map: Mapping[str, tuple[str, relativedelta | None]] = {
        "empty": ("", None),
        "zeros": ("P0Y0M0DT0H0M0S", None),
        "ones": (
            "P1Y1M1DT1H1M1S",
            relativedelta(years=1, months=1, days=1, hours=1, minutes=1, seconds=1),
        ),
        "years": ("P1YT", relativedelta(years=1)),
        "months": ("P1MT", relativedelta(months=1)),
        "days": ("P1DT", relativedelta(days=1)),
        "hours": ("PT1H", relativedelta(hours=1)),
        "minutes": ("PT1M", relativedelta(minutes=1)),
        "seconds": ("PT1S", relativedelta(seconds=1)),
    }

    @pytest.fixture(params=value_map.values(), ids=list(value_map.keys()))
    def duration_tuple(self, request) -> tuple[str, relativedelta | None]:  # noqa: ANN001
        """Duration tuple fixture."""
        return request.param

    def test_duration_tuple(self, duration_tuple: tuple[str, relativedelta | None]) -> None:
        """Test for from_duration_str()."""
        value: str = duration_tuple[0]
        expected: relativedelta | None = duration_tuple[1]

        obj: relativedelta | None = from_duration_str(value)

        assert isinstance(obj, relativedelta | None)
        assert obj == expected


class TestToDurationStr:
    """Tests for to_duration_str()."""

    value_map: Mapping[str, tuple[relativedelta | None, str]] = {
        "empty": (None, "PT0S"),
        "zeros": (relativedelta(years=0, months=0, days=0, hours=0, minutes=0, seconds=0), "PT0S"),
        "ones": (
            relativedelta(years=1, months=1, days=1, hours=1, minutes=1, seconds=1),
            "P1Y1M1DT1H1M1S",
        ),
        "years": (relativedelta(years=1), "P1YT"),
        "months": (relativedelta(months=1), "P1MT"),
        "days": (relativedelta(days=1), "P1DT"),
        "hours": (relativedelta(hours=1), "PT1H"),
        "minutes": (relativedelta(minutes=1), "PT1M"),
        "seconds": (relativedelta(seconds=1), "PT1S"),
    }

    @pytest.fixture(params=value_map.values(), ids=list(value_map.keys()))
    def duration_tuple(self, request) -> tuple[relativedelta | None, str]:  # noqa: ANN001
        """Duration tuple fixture."""
        return request.param

    def test_duration_tuple(self, duration_tuple: tuple[relativedelta | None, str]) -> None:
        """Test for to_duration_str()."""
        value: relativedelta | None = duration_tuple[0]
        expected: str = duration_tuple[1]

        obj: str = to_duration_str(value)

        assert isinstance(obj, str)
        assert obj == expected


if __name__ == "__main__":
    pytest.main()
