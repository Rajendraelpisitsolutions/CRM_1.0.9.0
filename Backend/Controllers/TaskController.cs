using Elpis_CRM.Model;
using Elpis_CRM.Services;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Authentication.JwtBearer;
using System.Threading.Tasks;

namespace Elpis_CRM.Controllers
{
    /// <summary>
    /// CRUD and reminder endpoints for CRM tasks. Most actions allow Admin, Manager, or User;
    /// deletion is restricted to Admin. Note the reminder endpoints are currently inert because
    /// reminders are handled client-side.
    /// </summary>
    [Route("api/[controller]")]
    [ApiController]
    public class TaskListController : ControllerBase
    {
        private readonly TaskService _taskService;

        /// <summary>
        /// Creates the controller with the task service used for persistence and queries.
        /// </summary>
        /// <param name="taskService">Service that stores and retrieves tasks.</param>
        public TaskListController(TaskService taskService)
        {
            _taskService = taskService;
        }

        /// <summary>
        /// Creates a task, stamping its created/updated timestamps before saving.
        /// </summary>
        /// <param name="task">Task payload to persist.</param>
        /// <returns>201 with the stored task and a Location header to its resource, or 400 if the body is missing.</returns>
        /// <response code="201">The task was created.</response>
        /// <response code="400">The request body was missing or could not be bound.</response>
        [HttpPost]
        [Authorize(AuthenticationSchemes = JwtBearerDefaults.AuthenticationScheme, Roles = "Admin,Manager,User")]
        public async Task<IActionResult> Add([FromBody] TaskModel task)
        {
            if (task == null)
            {
                return BadRequest("Invalid task data");
            }

            var created = await _taskService.AddTask(task);

            return CreatedAtAction(
                nameof(GetById),
                new { taskId = created.Id },
                created
            );
        }

        /// <summary>
        /// Returns every task in the system.
        /// </summary>
        /// <returns>200 with the list of tasks (empty list if none exist).</returns>
        /// <response code="200">Tasks retrieved.</response>
        [HttpGet]
        [Authorize(AuthenticationSchemes = JwtBearerDefaults.AuthenticationScheme, Roles = "Admin,Manager,User")]
        public async Task<IActionResult> GetAll()
        {
            var tasks = await _taskService.GetAllTasks();
            return Ok(tasks);
        }

        /// <summary>
        /// Fetches a single task by its identifier.
        /// </summary>
        /// <param name="taskId">Identifier of the task to retrieve.</param>
        /// <returns>200 with the task, or 404 if no task has that id.</returns>
        /// <response code="200">The task was found.</response>
        /// <response code="404">No task exists with the given id.</response>
        [HttpGet("{taskId:long}")]
        [Authorize(AuthenticationSchemes = JwtBearerDefaults.AuthenticationScheme, Roles = "Admin,Manager,User")]
        public async Task<IActionResult> GetById(int taskId)
        {
            var task = await _taskService.GetById(taskId);

            if (task == null)
            {
                return NotFound("Task not found");
            }

            return Ok(task);
        }

          /// <summary>
        /// Lists the tasks for a contact, ordered by due date then creation date (newest first).
        /// </summary>
        /// <param name="contactId">Identifier of the contact whose tasks are requested.</param>
        /// <returns>200 with the contact's tasks (empty list if none).</returns>
        /// <response code="200">Contact tasks retrieved.</response>
        [HttpGet("contact/{contactId:long}")]
        [Authorize(AuthenticationSchemes = JwtBearerDefaults.AuthenticationScheme, Roles = "Admin,Manager,User")]
        public async Task<IActionResult> GetTasksByContactId(long contactId)
        {
            var tasks = await _taskService.GetTasksByContactIdAsync(contactId);
            return Ok(tasks);
        }

        /// <summary>
        /// Replaces the editable fields of an existing task and refreshes its UpdatedAt timestamp.
        /// </summary>
        /// <param name="taskId">Identifier of the task to update.</param>
        /// <param name="task">Payload supplying the new field values.</param>
        /// <returns>200 with the updated task, or 404 if the task does not exist.</returns>
        /// <response code="200">The task was updated.</response>
        /// <response code="404">No task exists with the given id.</response>
        [HttpPut("{taskId:long}")]
        [Authorize(AuthenticationSchemes = JwtBearerDefaults.AuthenticationScheme, Roles = "Admin,Manager,User")]
        public async Task<IActionResult> Update(int taskId, [FromBody] TaskModel task)
        {
            var updated = await _taskService.UpdateTask(taskId, task);

            if (updated == null)
            {
                return NotFound("Task not found");
            }

            return Ok(updated);
        }

        /// <summary>
        /// Permanently removes a task. Admin-only.
        /// </summary>
        /// <param name="taskId">Identifier of the task to delete.</param>
        /// <returns>200 with a confirmation message, or 404 if the task does not exist.</returns>
        /// <response code="200">The task was deleted.</response>
        /// <response code="404">No task exists with the given id.</response>
        [HttpDelete("{taskId:long}")]
        [Authorize(AuthenticationSchemes = JwtBearerDefaults.AuthenticationScheme, Roles = "Admin")]
        public async Task<IActionResult> Delete(int taskId)
        {
            var deleted = await _taskService.DeleteTask(taskId);

            if (!deleted)
            {
                return NotFound("Task not found");
            }

            return Ok("Task deleted successfully");
        }

        /// <summary>
        /// Endpoint kept for API compatibility; reminders are computed client-side, so this always
        /// returns an empty list.
        /// </summary>
        /// <returns>200 with an empty task list.</returns>
        /// <response code="200">Always returns an empty list.</response>
        [HttpGet("due-reminders")]
        [Authorize(AuthenticationSchemes = JwtBearerDefaults.AuthenticationScheme, Roles = "Admin,Manager,User")]
        public async Task<IActionResult> GetDueReminders()
        {
            var reminders = await _taskService.GetDueReminders();
            return Ok(reminders);
        }

        /// <summary>
        /// Compatibility no-op for marking a reminder as sent; reminders are tracked client-side,
        /// so the service performs no work and reports success.
        /// </summary>
        /// <param name="taskId">Identifier of the task the reminder belongs to.</param>
        /// <returns>200 with a confirmation message (the underlying service always succeeds).</returns>
        /// <response code="200">Reminder acknowledged.</response>
        [HttpPost("{taskId:int}/mark-reminder-sent")]
        [Authorize(AuthenticationSchemes = JwtBearerDefaults.AuthenticationScheme, Roles = "Admin,Manager,User")]
        public async Task<IActionResult> MarkReminderAsSent(int taskId)
        {
            var result = await _taskService.MarkReminderAsSent(taskId);

            if (!result)
            {
                return NotFound("Task not found");
            }

            return Ok("Reminder marked as sent");
        }
    }
}