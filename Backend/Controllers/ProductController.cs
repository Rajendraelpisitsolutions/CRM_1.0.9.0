using Elpis_CRM.Model;
using Elpis_CRM.Services;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Authentication.JwtBearer;
using System;
using System.Threading.Tasks;

namespace Elpis_CRM.Controllers
{
    /// <summary>
    /// Controller for managing products in the CRM system.
    /// </summary>
    [ApiController]
    [Route("api/[controller]")]
    public class ProductsController : ControllerBase
    {
        private readonly ProductService _productService;
        private readonly RecycleBinService _recycleBinService;

        /// <summary>
        /// Initializes a new instance of the <see cref="ProductsController"/> class.
        /// </summary>
        /// <param name="productService">Service for product-related operations.</param>
        public ProductsController(ProductService productService, RecycleBinService recycleBinService)
        {
            _productService = productService;
            _recycleBinService = recycleBinService;
        }

        /// <summary>
        /// Returns the full product catalogue. Restricted to Admin, Manager and User roles.
        /// </summary>
        /// <returns>All products on success; on an unexpected error, a 500 carrying the exception message and stack trace (dev-only diagnostics).</returns>
        /// <response code="200">Products retrieved.</response>
        /// <response code="500">An unexpected error occurred; response includes exception detail.</response>
        [HttpGet]
        [Authorize(AuthenticationSchemes = JwtBearerDefaults.AuthenticationScheme, Roles = "Admin,Manager,User")]
        public async Task<IActionResult> GetAll()
        {
            try
            {
                var products = await _productService.GetAllAsync();
                return Ok(products);
            }
            catch (Exception ex)
            {
                // Development-time convenience: return exception details as JSON to help debugging.
                // Remove or restrict this in production.
                return StatusCode(500, new { message = ex.Message, detail = ex.ToString() });
            }
        }

        /// <summary>
        /// Looks up one product by primary key. Restricted to Admin, Manager and User roles.
        /// </summary>
        /// <param name="id">Primary key of the product.</param>
        /// <returns>The matching product, or a 404 when no product has that ID.</returns>
        /// <response code="200">Product found.</response>
        /// <response code="404">No product exists with the given ID.</response>
        [HttpGet("{id:int}")]
        [Authorize(AuthenticationSchemes = JwtBearerDefaults.AuthenticationScheme, Roles = "Admin,Manager,User")]
        public async Task<IActionResult> GetById(int id)
        {
            var product = await _productService.GetByIdAsync(id);
            if (product == null)
            {
                return NotFound($"Product with ID '{id}' not found.");
            }
            return Ok(product);
        }

        /// <summary>
        /// Finds the first product whose name matches exactly. Restricted to Admin, Manager and User roles.
        /// </summary>
        /// <param name="name">Exact product name to match.</param>
        /// <returns>The matching product, or a 404 when no product carries that name.</returns>
        /// <response code="200">Product found.</response>
        /// <response code="404">No product matches the given name.</response>
        [HttpGet("name/{name}")]
        [Authorize(AuthenticationSchemes = JwtBearerDefaults.AuthenticationScheme, Roles = "Admin,Manager,User")]
        public async Task<IActionResult> GetByName(string name)
        {
            var product = await _productService.GetProductByNameAsync(name);
            if (product == null)
            {
                return NotFound($"Product with name '{name}' not found.");
            }
            return Ok(product);
        }

        /// <summary>
        /// Returns every product in the given category. Restricted to Admin, Manager and User roles.
        /// </summary>
        /// <param name="category">Exact category name to filter on.</param>
        /// <returns>The products in that category, wrapped in 200 OK (empty list if none match).</returns>
        /// <response code="200">Matching products returned (possibly an empty list).</response>
        [HttpGet("category/{category}")]
        [Authorize(AuthenticationSchemes = JwtBearerDefaults.AuthenticationScheme, Roles = "Admin,Manager,User")]
        public async Task<IActionResult> GetByCategory(string category)
        {
            var products = await _productService.GetByCategoryAsync(category);
            return Ok(products);
        }

        /// <summary>
        /// Creates a product after null and model-state validation; the service stamps the timestamps.
        /// Restricted to Admin, Manager and User roles.
        /// </summary>
        /// <param name="product">Product payload to create.</param>
        /// <returns>The created product on success; a 400 for a null or invalid payload; a 500 with exception detail on an unexpected error.</returns>
        /// <response code="201">Product created; a Location header points to the new resource.</response>
        /// <response code="400">Payload was null or failed model validation.</response>
        /// <response code="500">An unexpected error occurred; response includes exception detail (dev-only).</response>
        [HttpPost]
        [Authorize(AuthenticationSchemes = JwtBearerDefaults.AuthenticationScheme, Roles = "Admin,Manager,User")]
        public async Task<IActionResult> Add([FromBody] ProductsModel product)
        {
            if (product == null)
            {
                return BadRequest("Product payload is null");
            }

            if (!ModelState.IsValid)
            {
                return BadRequest(ModelState);
            }

            try
            {
                var added = await _productService.AddAsync(product);
                return CreatedAtAction(nameof(GetById), new { id = added.ProductId }, added);
            }
            catch (Exception ex)
            {
                // Dev-only: return exception details to help debugging
                return StatusCode(500, new { message = ex.Message, detail = ex.ToString() });
            }
        }

        /// <summary>
        /// Updates a product's mutable fields. Restricted to Admin and Manager roles (User cannot update).
        /// </summary>
        /// <param name="id">Primary key of the product to update.</param>
        /// <param name="product">Payload carrying the new field values.</param>
        /// <returns>The updated product, or a 404 when the ID is unknown.</returns>
        /// <response code="200">Product updated.</response>
        /// <response code="404">No product exists with the given ID.</response>
        [HttpPut("{id:int}")]
        [Authorize(AuthenticationSchemes = JwtBearerDefaults.AuthenticationScheme, Roles = "Admin,Manager")]
        public async Task<IActionResult> Update(int id, [FromBody] ProductsModel product)
        {
            var updated = await _productService.UpdateAsync(id, product);
            if (updated == null)
            {
                return NotFound($"Product with ID '{id}' not found.");
            }
            return Ok(updated);
        }

        /// <summary>
        /// Permanently removes a product. Restricted to the Admin role only.
        /// </summary>
        /// <param name="id">Primary key of the product to delete.</param>
        /// <returns>A plain "Deleted Successfully" message, or a 404 when the ID is unknown.</returns>
        /// <response code="200">Product deleted.</response>
        /// <response code="404">No product exists with the given ID.</response>
        [HttpDelete("{id:int}")]
        [Authorize(AuthenticationSchemes = JwtBearerDefaults.AuthenticationScheme, Roles = "Admin")]
        public async Task<IActionResult> Delete(int id)
        {
            var product = await _productService.GetByIdAsync(id);
            if (product == null)
            {
                return NotFound($"Product with ID '{id}' not found.");
            }

            var deleted = await _productService.DeleteAsync(id, User?.Identity?.Name ?? User?.FindFirst(System.Security.Claims.ClaimTypes.Email)?.Value ?? "System");
            if (!deleted)
            {
                return NotFound($"Product with ID '{id}' not found.");
            }

            await _recycleBinService.CreateEntryAsync("Product", id.ToString(), product.Name ?? "Unnamed Product", "Product deleted", User?.Identity?.Name ?? User?.FindFirst(System.Security.Claims.ClaimTypes.Email)?.Value ?? "System", product);
            return Ok("Deleted Successfully");
        }
    }
}
