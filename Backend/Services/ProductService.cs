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
    /// Service class to handle CRUD operations for products.
    /// </summary>
    public class ProductService
    {
        private readonly AppDbContext _productDb;

        /// <summary>
        /// Initializes a new instance of the <see cref="ProductService"/> class.
        /// </summary>
        /// <param name="productDb">The database context for products.</param>
        public ProductService(AppDbContext productDb)
        {
            _productDb = productDb;
        }

        /// <summary>
        /// Loads every product, both active and inactive, in no particular order.
        /// </summary>
        /// <returns>All products as a list; empty when the table holds no rows.</returns>
        public async Task<List<ProductsModel>> GetAllAsync()
        {
            var productList = await _productDb.Products.ToListAsync();
            return productList;
        }

        /// <summary>
        /// Fetches a product by primary key using the context's identity-map-aware Find.
        /// </summary>
        /// <param name="productId">Primary key to look up.</param>
        /// <returns>The matching product, or null when no row has that ID.</returns>
        public async Task<ProductsModel?> GetByIdAsync(int productId)
        {
            return await _productDb.Products.FindAsync(productId);
        }

        /// <summary>
        /// Returns the first product whose Name equals the given value.
        /// </summary>
        /// <param name="name">Exact name to match.</param>
        /// <returns>The first matching product, or null when none match.</returns>
        public async Task<ProductsModel?> GetProductByNameAsync(string name)
        {
            return await _productDb.Products
                .FirstOrDefaultAsync(p => p.Name == name);
        }

        /// <summary>
        /// Loads every product whose Category exactly equals the given value.
        /// </summary>
        /// <param name="category">Category name to filter on.</param>
        /// <returns>The matching products; empty when none share that category.</returns>
        public async Task<List<ProductsModel>> GetByCategoryAsync(string category)
        {
            var productCategory = await _productDb.Products
                                   .Where(p => p.Category == category)
                                   .ToListAsync();
            return productCategory;
        }

        /// <summary>
        /// Inserts a product, setting both CreatedAt and UpdatedAt to the current UTC time before saving.
        /// </summary>
        /// <param name="product">The product to persist.</param>
        /// <returns>The same instance after saving, now carrying its database-generated ID and timestamps.</returns>
        public async Task<ProductsModel> AddAsync(ProductsModel product)
        {
            product.CreatedAt = DateTime.UtcNow;
            product.UpdatedAt = DateTime.UtcNow;

            _productDb.Products.Add(product);
            await _productDb.SaveChangesAsync();
            return product;
        }

        /// <summary>
        /// Copies name, active flag, base-currency amount, category and UpdatedBy onto the existing row and bumps
        /// UpdatedAt to now (UTC). CreatedAt, the ID and any other columns are left untouched.
        /// </summary>
        /// <param name="productId">Primary key of the product to update.</param>
        /// <param name="product">Source of the new field values.</param>
        /// <returns>The updated product, or null when no row has that ID.</returns>
        public async Task<ProductsModel?> UpdateAsync(int productId, ProductsModel product)
        {
            var existing = await _productDb.Products.FindAsync(productId);
            if (existing == null)
            {
                return null;
            }

            existing.Name = product.Name;
            existing.Active = product.Active;
            existing.BaseCurrencyAmount = product.BaseCurrencyAmount;
            existing.Category = product.Category;
            existing.UpdatedAt = DateTime.UtcNow;
            existing.UpdatedBy = product.UpdatedBy;

            await _productDb.SaveChangesAsync();
            return existing;
        }

        /// <summary>
        /// Hard-deletes the product with the given ID, if it exists.
        /// </summary>
        /// <param name="productId">Primary key of the product to remove.</param>
        /// <returns>True when a row was found and deleted; false when the ID was not present.</returns>
        public async Task<bool> DeleteAsync(int productId)
        {
            var product = await _productDb.Products.FindAsync(productId);

            if (product == null)
            {
                return false;
            }
            _productDb.Products.Remove(product);
            await _productDb.SaveChangesAsync();
            return true;
        }
    }
}
