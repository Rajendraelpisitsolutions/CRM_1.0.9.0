import React, { useState } from "react";
import { Filter, Plus, Search } from "lucide-react";
import FilterItem from "./FilterItem";

function FilterBuilder({
  filters,
  columns,
  tagOptions,
  selectedTags,
  onAddFilter,
  onRemoveFilter,
  onFilterChange,
  onTagsChange,
  getFilterableFields,
  getFieldType,
  getOperatorsByType,
  allowAddFilter = true,
}) {
  const [tagSearch, setTagSearch] = useState("");
  if (
    typeof tagOptions !== "undefined" &&
    tagOptions &&
    tagOptions.length > 0
  ) {
    const sel = Array.isArray(selectedTags) ? selectedTags : [];
    const toggle = (t) =>
      onTagsChange(sel.includes(t) ? sel.filter((x) => x !== t) : [...sel, t]);
    const q = tagSearch.trim().toLowerCase();
    const shown = (Array.isArray(tagOptions) ? tagOptions : []).filter(
      (t) => !q || String(t).toLowerCase().includes(q)
    );
    return (
      <div className="space-y-3">
        <div className="flex items-center justify-between gap-2">
          <div className="text-sm text-gray-600">
            Select tags to filter <span className="text-gray-400">(matches any)</span>
          </div>
          {sel.length > 0 && (
            <button
              type="button"
              onClick={() => onTagsChange([])}
              className="text-xs font-medium text-blue-600 hover:text-blue-800 whitespace-nowrap"
            >
              Clear ({sel.length})
            </button>
          )}
        </div>
        {/* Search — find a tag instead of scrolling the whole list */}
        <div className="relative">
          <Search className="w-4 h-4 text-gray-400 absolute left-3 top-1/2 -translate-y-1/2" />
          <input
            type="text"
            value={tagSearch}
            onChange={(e) => setTagSearch(e.target.value)}
            placeholder="Search tags…"
            className="w-full pl-9 pr-3 py-2 text-sm border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent outline-none"
          />
        </div>
        <div className="max-h-60 overflow-y-auto border border-gray-200 rounded-lg divide-y divide-gray-100">
          {shown.length > 0 ? (
            shown.map((t) => (
              <label
                key={t}
                className="flex items-center gap-2.5 px-3 py-2 hover:bg-gray-50 cursor-pointer text-sm select-none"
              >
                <input
                  type="checkbox"
                  checked={sel.includes(t)}
                  onChange={() => toggle(t)}
                  className="w-4 h-4 rounded border-gray-300 text-blue-600 focus:ring-2 focus:ring-blue-500"
                />
                <span className="text-gray-700 truncate">{t}</span>
              </label>
            ))
          ) : (
            <div className="px-3 py-4 text-sm text-gray-400 text-center">
              {q ? "No matching tags" : "No tags available"}
            </div>
          )}
        </div>
      </div>
    );
  }

  if (filters.length === 0 && !allowAddFilter) {
    return null;
  }

  if (filters.length === 0) {
    return (
      <div className="text-center py-8">
        <Filter className="w-12 h-12 text-gray-300 mx-auto mb-3" />
        <p className="text-gray-500 text-sm mb-4">
          No filters applied
        </p>
        {allowAddFilter && (
          <button
            onClick={onAddFilter}
            className="inline-flex items-center gap-2 px-4 py-2 bg-blue-500 hover:bg-blue-600 text-white rounded-lg text-sm font-medium transition-colors"
          >
            <Plus className="w-4 h-4" />
            Add Filter
          </button>
        )}
      </div>
    );
  }

  return (
    <>
      {filters.map((filter, index) => (
        <FilterItem
          key={filter.id}
          filter={filter}
          index={index}
          columns={columns}
          onFieldChange={(id, field) =>
            onFilterChange(id, "field", field)
          }
          onOperatorChange={(id, operator) =>
            onFilterChange(id, "operator", operator)
          }
          onValueChange={(id, value) =>
            onFilterChange(id, "value", value)
          }
          onRemove={onRemoveFilter}
          getFilterableFields={getFilterableFields}
          getFieldType={getFieldType}
          getOperatorsByType={getOperatorsByType}
        />
      ))}

      {allowAddFilter && (
        <button
          onClick={onAddFilter}
          className="w-full inline-flex items-center justify-center gap-2 px-4 py-2 bg-white hover:bg-gray-50 border-2 border-dashed border-gray-300 hover:border-blue-400 text-gray-700 hover:text-blue-700 rounded-lg text-sm font-medium transition-all"
        >
          <Plus className="w-4 h-4" />
          Add Another Filter
        </button>
      )}
    </>
  );
}

export default FilterBuilder;