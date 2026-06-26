using Elpis_CRM.Data;
using Elpis_CRM.Model;
using Microsoft.EntityFrameworkCore;
using System.Net.Mail;
using System.Net;
using System.Security.Cryptography;
using System.Text;
using System.Collections.Concurrent;

/// <summary>
/// Drives the forgot-password flow: issues time-limited OTPs, verifies them, and resets the stored password.
/// OTPs are held in a process-wide in-memory store keyed by email, so they do not survive a restart and are
/// not shared across instances.
/// </summary>
public class ForgotPasswordService
{
    private readonly AppDbContext _loginDb;
    private static readonly ConcurrentDictionary<string, OtpModel> otpStore = new();

    public ForgotPasswordService(AppDbContext loginDb)
    {
        _loginDb = loginDb;
    }

    /// <summary>
    /// Produces a cryptographically-random 4-digit OTP, records it against the email with a 5-minute expiry
    /// (overwriting any prior code), and emails it. No account-existence check is performed.
    /// </summary>
    /// <param name="email">Address that will receive the code and act as the store key.</param>
    /// <returns>"OTP sent successfully", or "Email cannot be empty" when the address is blank.</returns>
    public async Task<string> GenerateOtpAsync(string email)
    {
        if (string.IsNullOrWhiteSpace(email))
        {
            return "Email cannot be empty";
        }

        int otp = RandomNumberGenerator.GetInt32(1000, 10000);

        otpStore[email] = new OtpModel
        {
            Email = email,
            Otp = otp,
            ExpiryTime = DateTime.UtcNow.AddMinutes(5)
        };

        await SendEmailAsync(email, otp);
        return "OTP sent successfully";
    }

    /// <summary>
    /// Confirms the supplied code matches the unexpired OTP stored for the email; on success it overwrites the
    /// stored code with the sentinel -1 to mark it verified for the password-reset step.
    /// </summary>
    /// <param name="email">Address whose pending OTP is being checked.</param>
    /// <param name="otp">The code entered by the user.</param>
    /// <returns>"OTP validated successfully", "OTP not found", or "Invalid or expired OTP".</returns>
    public string ValidateOtp(string email, int otp)
    {
        if (!otpStore.TryGetValue(email, out var data))
        {
            return "OTP not found";
        }
        if (data.Otp == otp && data.ExpiryTime > DateTime.UtcNow)
        {
            // Mark as validated
            otpStore[email].Otp = -1;
            return "OTP validated successfully";
        }

        return "Invalid or expired OTP";
    }

    /// <summary>
    /// Writes the new password for the account, but only when the email carries a validated OTP (the -1 sentinel);
    /// it then removes the OTP entry so the same validation cannot be reused.
    /// </summary>
    /// <param name="email">Address identifying the account; must have a validated OTP on file.</param>
    /// <param name="newPassword">Replacement password, stored as provided.</param>
    /// <returns>"Password updated successfully", "OTP not validated", or "User not found".</returns>
    public async Task<string> UpdatePasswordAsync(string email, string newPassword)
    {
        if (!otpStore.TryGetValue(email, out var otpEntry) || otpEntry.Otp != -1)
        {
            return "OTP not validated";
        }
        var user = await _loginDb.Login.FirstOrDefaultAsync(u => u.Email == email);
        if (user == null)
        {
            return "User not found";
        }
        user.Password = newPassword;
        await _loginDb.SaveChangesAsync();

        otpStore.TryRemove(email, out _); 
        return "Password updated successfully";
    }

    #region Helpers

    /// <summary>
    /// Sends the OTP to the recipient over the configured Office 365 SMTP relay (TLS on port 587).
    /// </summary>
    /// <param name="userEmail">Recipient address; trimmed before use.</param>
    /// <param name="otp">The one-time code to include in the message body.</param>
    private async Task SendEmailAsync(string userEmail, int otp)
    {
        var fromAddress = new MailAddress("Rajendra.m@elpisitsolutions.com", "Elpis CRM OTP Service");
        var toAddress = new MailAddress(userEmail.Trim());
        string fromPassword = "Rajendra123@@";

        using var smtp = new SmtpClient("smtp.office365.com", 587)
        {
            EnableSsl = true,
            Credentials = new NetworkCredential(fromAddress.Address, fromPassword)
        };

        using var message = new MailMessage(fromAddress, toAddress)
        {
            Subject = "Your OTP Code",
            Body = $"Your OTP code is {otp}. It will expire in 5 minutes."
        };

        await smtp.SendMailAsync(message);
    }
    #endregion
}
