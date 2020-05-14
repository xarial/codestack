public class TaskPaneControl : UserControl
{
}
public enum TaskPaneCommands_e
{
    Command1
}
    //...
    TaskPaneControl ctrl;
    var taskPaneView = CreateTaskPane<TaskPaneControl, TaskPaneCommands_e>(OnTaskPaneCommandClick, out ctrl);
    //...
private void OnTaskPaneCommandClick(TaskPaneCommands_e cmd)
{
    switch (cmd)
    {
        case TaskPaneCommands_e.Command1:
            //TODO: handle command
            break;
    }
}