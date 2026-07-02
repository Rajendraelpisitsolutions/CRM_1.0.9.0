import React, { useState, useEffect, useCallback } from "react";
import { useNavigate } from "react-router-dom";
import apiClient from "../api/client";
import SignupPanel from "./SignupPanel";
import EditUserModal from "../components/modals/EditUserModal";
import PasswordModal from "../components/modals/PasswordModal";
import DeleteConfirmModal from "../components/modals/DeleteConfirmModal";
import {
  FiEdit2,
  FiTrash2,
  FiPlus,
  FiSearch,
  FiCheck,
  FiX,
  FiLock,
} from "react-icons/fi";

const STATUS_STYLES = {
  Active: "bg-green-100 text-green-700 border-green-200",
  Inactive: "bg-red-100 text-red-700 border-red-200",
};

const ROLE_STYLES = {
  Admin: "bg-purple-100 text-purple-700",
  Manager: "bg-blue-100 text-blue-700",
  User: "bg-gray-100 text-gray-700",
};

function Users() {
  const navigate = useNavigate();

  // Users Data & State
  const [rows, setRows] = useState([]);
  const [total, setTotal] = useState(0);
  const [page, setPage] = useState(1);
  const [pageSize] = useState(50);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState(null);

  // Filters
  const [search, setSearch] = useState("");
  const [roleFilter, setRoleFilter] = useState("");
  const [statusFilter, setStatusFilter] = useState("");
  const [roles, setRoles] = useState([]);

  // Modal State
  const [showSignupPanel, setShowSignupPanel] = useState(false);
  const [showEditModal, setShowEditModal] = useState(false);
  const [showPasswordModal, setShowPasswordModal] = useState(false);
  const [showDeleteConfirm, setShowDeleteConfirm] = useState(false);
  const [successMessage, setSuccessMessage] = useState("");

  // Edit & Delete State
  const [editData, setEditData] = useState({
    loginId: null,
    name: "",
    email: "",
    phone: "",
    role: "User",
    isActive: true,
  });

  const [passwordData, setPasswordData] = useState({
    loginId: null,
    newPassword: "",
    confirmPassword: "",
  });

  const [deleteId, setDeleteId] = useState(null);
  const [errors, setErrors] = useState({});

  // Fetch users with filters
  const fetchUsers = useCallback(async () => {
    setLoading(true);
    setError(null);
    try {
      const response = await apiClient.get("/Login/all");
      let users = Array.isArray(response.data) ? response.data : [];

      // Apply filters
      if (search) {
        const searchLower = search.toLowerCase();
        users = users.filter(
          (u) =>
            (u.name?.toLowerCase().includes(searchLower)) ||
            (u.email?.toLowerCase().includes(searchLower)) ||
            (u.phone?.includes(search))
        );
      }

      if (roleFilter) {
        users = users.filter((u) => u.role === roleFilter);
      }

      if (statusFilter) {
        const isActive = statusFilter === "Active";
        users = users.filter((u) => u.isActive === isActive);
      }

      // Sort by name
      users.sort((a, b) => (a.name || "").localeCompare(b.name || ""));

      setTotal(users.length);

      // Apply pagination
      const start = (page - 1) * pageSize;
      const paginated = users.slice(start, start + pageSize);
      setRows(paginated);
      setErrors({});
    } catch (err) {
      console.error("User fetch failed", err);
      setError(err?.response?.data?.message || err?.message || "Failed to load users.");
      setRows([]);
      setTotal(0);
    } finally {
      setLoading(false);
    }
  }, [search, roleFilter, statusFilter, page, pageSize]);

  // Extract unique roles for filter dropdown
  const extractRoles = useCallback(async () => {
    try {
      const response = await apiClient.get("/Login/all");
      const users = Array.isArray(response.data) ? response.data : [];
      const uniqueRoles = [...new Set(users.map((u) => u.role).filter(Boolean))];
      setRoles(uniqueRoles.sort());
    } catch (err) {
      console.error("Failed to extract roles", err);
      setRoles([]);
    }
  }, []);

  useEffect(() => {
    extractRoles();
  }, [extractRoles]);

  useEffect(() => {
    fetchUsers();
  }, [fetchUsers]);

  const totalPages = Math.max(1, Math.ceil(total / pageSize));

  const handleOpenForm = (user = null) => {
    if (user) {
      // Open edit modal with user data
      setEditData({
        loginId: user.loginId,
        name: user.name || "",
        email: user.email || "",
        phone: user.phone || "",
        role: user.role || "User",
        isActive: user.isActive ?? true,
      });
      setShowEditModal(true);
    } else {
      // Open signup panel for new user
      setShowSignupPanel(true);
    }
    setErrors({});
  };

  const handleCloseForm = () => {
    setShowSignupPanel(false);
    setShowEditModal(false);
    setEditData({
      loginId: null,
      name: "",
      email: "",
      phone: "",
      role: "User",
      isActive: true,
    });
    setErrors({});
  };

  const handleEditChange = (e) => {
    const { name, value, type, checked } = e.target;
    setEditData({
      ...editData,
      [name]: type === "checkbox" ? checked : value,
    });
  };

  const handleEditSubmit = async (e) => {
    e.preventDefault();

    // Validation
    if (!editData.name.trim()) {
      setErrors({ name: "Name is required" });
      return;
    }
    if (!editData.email.trim()) {
      setErrors({ email: "Email is required" });
      return;
    }
    if (!editData.phone.trim()) {
      setErrors({ phone: "Phone is required" });
      return;
    }

    try {
      const payload = {
        name: editData.name,
        email: editData.email,
        phone: editData.phone,
        role: editData.role,
        isActive: editData.isActive,
        password: "", // Send empty password to satisfy validation
      };

      // Ensure loginId is converted to integer
      const loginId = parseInt(editData.loginId, 10);
      
      console.log("Sending payload:", payload);
      console.log("LoginId:", loginId);
      
      await apiClient.put(`/Login/${loginId}`, payload);
      setSuccessMessage("User updated successfully!");
      setErrors({});

      setTimeout(() => {
        fetchUsers();
        handleCloseForm();
        setSuccessMessage("");
      }, 1500);
    } catch (error) {
      console.error("Edit error:", error);
      console.error("Error response data:", error?.response?.data);
      setErrors({ general: error?.response?.data?.message || JSON.stringify(error?.response?.data) || "Failed to update user" });
    }
  };

  const handleActivate = async (userId) => {
    try {
      await apiClient.put(`/Login/${userId}/activate`);
      setSuccessMessage("User activated successfully!");
      setTimeout(() => {
        fetchUsers();
        setSuccessMessage("");
      }, 1500);
    } catch (error) {
      setErrors({ general: "Failed to activate user" });
    }
  };

  const handleDeactivate = async (userId) => {
    try {
      await apiClient.put(`/Login/${userId}/deactivate`);
      setSuccessMessage("User deactivated successfully!");
      setTimeout(() => {
        fetchUsers();
        setSuccessMessage("");
      }, 1500);
    } catch (error) {
      setErrors({ general: "Failed to deactivate user" });
    }
  };

  const handleDelete = async (userId) => {
    try {
      await apiClient.delete(`/Login/${userId}`);
      setSuccessMessage("User deleted successfully!");
      setTimeout(() => {
        fetchUsers();
        setShowDeleteConfirm(false);
        setDeleteId(null);
        setSuccessMessage("");
      }, 1500);
    } catch (error) {
      setErrors({ general: "Failed to delete user" });
    }
  };

  const handleOpenPasswordModal = (userId) => {
    setPasswordData({
      loginId: userId,
      newPassword: "",
      confirmPassword: "",
    });
    setErrors({});
    setShowPasswordModal(true);
  };

  const handleChangePassword = async (e) => {
    e.preventDefault();

    if (!passwordData.newPassword.trim()) {
      setErrors({ password: "New password is required" });
      return;
    }

    if (passwordData.newPassword.length < 6) {
      setErrors({ password: "Password must be at least 6 characters" });
      return;
    }

    if (passwordData.newPassword !== passwordData.confirmPassword) {
      setErrors({ password: "Passwords do not match" });
      return;
    }

    try {
      // Use the change-password endpoint which gets email from JWT token
      await apiClient.put(`/Login/change-password`, {
        NewPassword: passwordData.newPassword,
      });

      setSuccessMessage("Password changed successfully!");
      setErrors({});
      setTimeout(() => {
        fetchUsers();
        setShowPasswordModal(false);
        setPasswordData({ loginId: null, newPassword: "", confirmPassword: "" });
        setSuccessMessage("");
      }, 1500);
    } catch (error) {
      setErrors({ general: error?.response?.data?.message || "Failed to change password" });
    }
  };

  return (
    <div className="p-4 sm:p-6 h-full overflow-hidden bg-gray-50">
      <div className="max-w-7xl mx-auto h-full flex flex-col min-h-0">
        <div className="sticky top-0 z-10 bg-gray-50 pb-3">
          {/* Back Button */}
          <button
            onClick={() => navigate(-1)}
            className="inline-flex items-center gap-1.5 text-sm text-gray-600 hover:text-gray-900 font-medium mb-3"
          >
            <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M15 19l-7-7 7-7" />
            </svg>
            Back
          </button>

          {/* Header */}
          <div className="flex items-center justify-between flex-wrap gap-3 mb-4">
            <div>
              <h1 className="text-xl font-bold text-gray-900">User Management</h1>
              <p className="text-sm text-gray-500">Manage users, roles, and permissions</p>
            </div>
            <div className="flex items-center gap-3">
              <span className="text-sm text-gray-500">{total} user{total !== 1 ? "s" : ""}</span>
              <button
                onClick={() => handleOpenForm()}
                className="flex items-center gap-2 px-4 py-2 bg-indigo-600 hover:bg-indigo-700 text-white text-sm font-medium rounded-lg"
              >
                <FiPlus size={16} />
                Add User
              </button>
            </div>
          </div>

          {/* Success Message */}
          {successMessage && (
            <div className="mb-4 px-4 py-3 rounded-lg bg-green-50 border border-green-200 text-green-700 text-sm flex items-center gap-3">
              <FiCheck size={18} />
              {successMessage}
            </div>
          )}

          {/* Error Message */}
          {error && (
            <div className="mb-4 px-4 py-3 rounded-lg bg-red-50 border border-red-200 text-red-700 text-sm flex items-center gap-3">
              <FiX size={18} />
              {error}
            </div>
          )}

          {/* Filters */}
          <div className="flex flex-wrap items-center gap-2 mb-4">
            <input
              value={search}
              onChange={(e) => {
                setSearch(e.target.value);
                setPage(1);
              }}
              placeholder="Search by name, email, or phone…"
              className="flex-1 min-w-[200px] px-3 py-2 border border-gray-300 rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500"
            />
            <select
              value={roleFilter}
              onChange={(e) => {
                setRoleFilter(e.target.value);
                setPage(1);
              }}
              className="px-3 py-2 border border-gray-300 rounded-lg text-sm bg-white focus:outline-none focus:ring-2 focus:ring-indigo-500"
            >
              <option value="">All roles</option>
              {roles.map((r) => (
                <option key={r} value={r}>
                  {r}
                </option>
              ))}
            </select>
            <select
              value={statusFilter}
              onChange={(e) => {
                setStatusFilter(e.target.value);
                setPage(1);
              }}
              className="px-3 py-2 border border-gray-300 rounded-lg text-sm bg-white focus:outline-none focus:ring-2 focus:ring-indigo-500"
            >
              <option value="">All statuses</option>
              <option value="Active">Active</option>
              <option value="Inactive">Inactive</option>
            </select>
            <button
              onClick={() => fetchUsers()}
              className="px-4 py-2 bg-indigo-600 hover:bg-indigo-700 text-white text-sm rounded-lg"
            >
              Refresh
            </button>
          </div>
        </div>

        {/* Table */}
        <div className="flex-1 min-h-0 overflow-auto bg-white border border-gray-200 rounded-xl">
          <div className="overflow-x-auto">
            <table className="w-full text-sm">
              <thead className="bg-gray-50 text-gray-500 text-xs uppercase tracking-wide sticky top-0 z-10">
                <tr>
                  <th className="text-left font-semibold px-4 py-3 whitespace-nowrap">Name</th>
                  <th className="text-left font-semibold px-4 py-3 whitespace-nowrap">Email</th>
                  <th className="text-left font-semibold px-4 py-3 whitespace-nowrap">Phone</th>
                  <th className="text-left font-semibold px-4 py-3 whitespace-nowrap">Role</th>
                  <th className="text-left font-semibold px-4 py-3 whitespace-nowrap">Status</th>
                  <th className="text-right font-semibold px-4 py-3 whitespace-nowrap">Actions</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-gray-100">
                {loading ? (
                  <tr>
                    <td colSpan={6} className="px-4 py-10 text-center text-gray-500">
                      Loading…
                    </td>
                  </tr>
                ) : rows.length === 0 ? (
                  <tr>
                    <td colSpan={6} className="px-4 py-10 text-center text-gray-400">
                      No users found.
                    </td>
                  </tr>
                ) : (
                  rows.map((u) => (
                    <tr key={u.loginId} className="hover:bg-gray-50 align-top">
                      <td className="px-4 py-3 whitespace-nowrap font-medium text-gray-800">{u.name}</td>
                      <td className="px-4 py-3 whitespace-nowrap text-gray-600">{u.email}</td>
                      <td className="px-4 py-3 whitespace-nowrap text-gray-600">{u.phone}</td>
                      <td className="px-4 py-3 whitespace-nowrap">
                        <span
                          className={`px-2 py-0.5 rounded-full text-xs font-medium ${
                            ROLE_STYLES[u.role] || "bg-gray-100 text-gray-700"
                          }`}
                        >
                          {u.role}
                        </span>
                      </td>
                      <td className="px-4 py-3 whitespace-nowrap">
                        <span
                          className={`px-2 py-0.5 rounded-full text-xs font-medium border ${
                            STATUS_STYLES[u.isActive ? "Active" : "Inactive"]
                          }`}
                        >
                          {u.isActive ? "Active" : "Inactive"}
                        </span>
                      </td>
                      <td className="px-4 py-3 text-right">
                        <div className="flex items-center justify-end gap-1">
                          <button
                            onClick={() => handleOpenForm(u)}
                            className="p-2 text-blue-600 hover:bg-blue-100 rounded transition"
                            title="Edit"
                          >
                            <FiEdit2 size={16} />
                          </button>
                          <button
                            onClick={() => handleOpenPasswordModal(u.loginId)}
                            className="p-2 text-orange-600 hover:bg-orange-100 rounded transition"
                            title="Change Password"
                          >
                            <FiLock size={16} />
                          </button>
                          {u.isActive ? (
                            <button
                              onClick={() => handleDeactivate(u.loginId)}
                              className="p-2 text-yellow-600 hover:bg-yellow-100 rounded transition"
                              title="Deactivate"
                            >
                              <FiX size={16} />
                            </button>
                          ) : (
                            <button
                              onClick={() => handleActivate(u.loginId)}
                              className="p-2 text-green-600 hover:bg-green-100 rounded transition"
                              title="Activate"
                            >
                              <FiCheck size={16} />
                            </button>
                          )}
                          <button
                            onClick={() => {
                              setDeleteId(u.loginId);
                              setShowDeleteConfirm(true);
                            }}
                            className="p-2 text-red-600 hover:bg-red-100 rounded transition"
                            title="Delete"
                          >
                            <FiTrash2 size={16} />
                          </button>
                        </div>
                      </td>
                    </tr>
                  ))
                )}
              </tbody>
            </table>
          </div>
        </div>

        {/* Pagination */}
        {totalPages > 1 && (
          <div className="mt-4 flex items-center justify-between text-sm text-gray-600">
            <span>
              Page {page} of {totalPages}
            </span>
            <div className="flex gap-2">
              <button
                onClick={() => setPage(Math.max(1, page - 1))}
                disabled={page === 1}
                className="px-3 py-1 border border-gray-300 rounded hover:bg-gray-50 disabled:opacity-50"
              >
                Previous
              </button>
              {[...Array(totalPages)].map((_, i) => (
                <button
                  key={i + 1}
                  onClick={() => setPage(i + 1)}
                  className={`px-3 py-1 rounded border ${
                    page === i + 1
                      ? "bg-indigo-600 text-white border-indigo-600"
                      : "border-gray-300 hover:bg-gray-50"
                  }`}
                >
                  {i + 1}
                </button>
              ))}
              <button
                onClick={() => setPage(Math.min(totalPages, page + 1))}
                disabled={page === totalPages}
                className="px-3 py-1 border border-gray-300 rounded hover:bg-gray-50 disabled:opacity-50"
              >
                Next
              </button>
            </div>
          </div>
        )}
      </div>

      {/* Modals */}
      {showSignupPanel && (
        <SignupPanel
          onClose={() => {
            setShowSignupPanel(false);
            fetchUsers();
          }}
        />
      )}

      <EditUserModal
        isOpen={showEditModal}
        editData={editData}
        errors={errors}
        successMessage={successMessage}
        onClose={handleCloseForm}
        onChange={handleEditChange}
        onSubmit={handleEditSubmit}
      />

      <PasswordModal
        isOpen={showPasswordModal}
        passwordData={passwordData}
        errors={errors}
        successMessage={successMessage}
        onClose={() => setShowPasswordModal(false)}
        onChange={(e) => {
          const { name, value } = e.target;
          setPasswordData({ ...passwordData, [name]: value });
        }}
        onSubmit={handleChangePassword}
      />

      <DeleteConfirmModal
        isOpen={showDeleteConfirm}
        onConfirm={() => handleDelete(deleteId)}
        onCancel={() => {
          setShowDeleteConfirm(false);
          setDeleteId(null);
        }}
      />
    </div>
  );
}

export default Users;
