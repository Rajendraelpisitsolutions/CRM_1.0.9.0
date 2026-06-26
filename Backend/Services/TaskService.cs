using Elpis_CRM.Data;
using Elpis_CRM.Model;
using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Elpis_CRM.Services
{
    /// <summary>
    /// Persistence and query logic for CRM tasks. The reminder methods are deliberately inert
    /// because reminder handling lives on the client.
    /// </summary>
    public class TaskService
    {
        private readonly AppDbContext _taskDb;

        /// <summary>
        /// Creates the service over the given database context.
        /// </summary>
        /// <param name="taskDb">EF Core context backing the Tasks table.</param>
        public TaskService(AppDbContext taskDb)
        {
            _taskDb = taskDb;
        }

        // Add Task
        /// <summary>
        /// Inserts a new task, setting its CreatedAt and UpdatedAt to the current UTC time.
        /// </summary>
        /// <param name="task">Task to persist.</param>
        /// <returns>The saved task, including its generated id.</returns>
        public async Task<TaskModel> AddTask(TaskModel task)
        {
            task.CreatedAt = DateTime.UtcNow;
            task.UpdatedAt = DateTime.UtcNow;
           

            _taskDb.Tasks.Add(task);
            await _taskDb.SaveChangesAsync();

            return task;
        }

        // GET ALL TASKS
        /// <summary>
        /// Returns all tasks in their natural store order.
        /// </summary>
        /// <returns>Every task, or an empty list if none exist.</returns>
        public async Task<List<TaskModel>> GetAllTasks()
        {
            return await _taskDb.Tasks.ToListAsync();
        }

        // GET TASK BY ID
        /// <summary>
        /// Looks up a single task by its primary key.
        /// </summary>
        /// <param name="id">Identifier of the task.</param>
        /// <returns>The matching task, or null if it is not found.</returns>
        public async Task<TaskModel?> GetById(int id)
        {
            return await _taskDb.Tasks.FindAsync(id);
        }

        // GET TASKS BY CONTACT ID
        /// <summary>
        /// Returns a contact's tasks (read-only), ordered by due date then creation time, both
        /// descending; null dates sort as DateTime.MinValue so they fall last.
        /// </summary>
        /// <param name="contactId">Identifier of the contact to filter by.</param>
        /// <returns>The contact's tasks, or an empty list if there are none.</returns>
        public async Task<List<TaskModel>> GetTasksByContactIdAsync(long contactId)
        {
            return await _taskDb.Tasks
                .AsNoTracking()
                .Where(t => t.ContactId == contactId)
                .OrderByDescending(t => t.DueDate ?? DateTime.MinValue)
                .ThenByDescending(t => t.CreatedAt ?? DateTime.MinValue)
                .ToListAsync();
        }
        
        // UPDATE TASK
        /// <summary>
        /// Copies the editable fields from the supplied task onto the stored one and bumps UpdatedAt
        /// to the current UTC time.
        /// </summary>
        /// <param name="id">Identifier of the task to update.</param>
        /// <param name="task">Source of the new field values.</param>
        /// <returns>The updated task, or null if no task has the given id.</returns>
        public async Task<TaskModel?> UpdateTask(int id, TaskModel task)
        {
            var existing = await _taskDb.Tasks.FindAsync(id);

            if (existing == null)
            {
                return null;
            }
            existing.Title = task.Title;
            existing.Description = task.Description;
            existing.Status = task.Status;
            existing.TaskType = task.TaskType;
            existing.DueDate = task.DueDate;
            existing.CompletedDate = task.CompletedDate;
            existing.Outcome = task.Outcome;
            existing.OwnerId = task.OwnerId;
            existing.CreatedById = task.CreatedById;
            existing.UpdatedById = task.UpdatedById;
            existing.UpdatedAt = DateTime.UtcNow;

            await _taskDb.SaveChangesAsync();
            return existing;
        }

        // DELETE TASK
        /// <summary>
        /// Removes a task by id.
        /// </summary>
        /// <param name="id">Identifier of the task to delete.</param>
        /// <returns>True if the task was found and removed; false if no task has the given id.</returns>
        public async Task<bool> DeleteTask(int id)
        {
            var task = await _taskDb.Tasks.FindAsync(id);

            if (task == null)
            {
                return false;
            }
            _taskDb.Tasks.Remove(task);
            await _taskDb.SaveChangesAsync();

            return true;
        }

        // GET TASKS DUE TODAY
        /// <summary>
        /// Returns tasks whose due date falls on the current local calendar day.
        /// </summary>
        /// <returns>Tasks due today, or an empty list if none qualify.</returns>
        public async Task<List<TaskModel>> GetTasksDueToday()
        {
            DateTime today = DateTime.Today;

            return await _taskDb.Tasks
                .Where(t => t.DueDate.HasValue && t.DueDate.Value.Date == today)
                .ToListAsync();
        }

        // GET TASKS WITH DUE REMINDERS - Returns empty (reminders handled client-side)
        /// <summary>
        /// Placeholder retained for API compatibility; reminders are evaluated client-side, so this
        /// always yields an empty list and queries nothing.
        /// </summary>
        /// <returns>An empty task list.</returns>
        public async Task<List<TaskModel>> GetDueReminders()
        {
            // Reminder functionality is handled client-side in useTaskReminders hook
            // Return empty list to maintain API compatibility
            return await Task.FromResult(new List<TaskModel>());
        }

        // MARK REMINDER AS SENT - No-op (reminders handled client-side)
        /// <summary>
        /// No-op kept for API compatibility; reminder state is owned by the client, so this performs
        /// no work and unconditionally reports success.
        /// </summary>
        /// <param name="id">Identifier of the task whose reminder would be marked (unused).</param>
        /// <returns>Always true.</returns>
        public async Task<bool> MarkReminderAsSent(int id)
        {
            // Reminder functionality is handled client-side
            // Return true to indicate successful processing
            return await Task.FromResult(true);
        }
    }
}