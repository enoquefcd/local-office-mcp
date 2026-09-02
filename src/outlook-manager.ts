import { spawn } from 'child_process';
import { writeFileSync, mkdtempSync, rmSync } from 'fs';
import { tmpdir } from 'os';
import { join } from 'path';

export interface EmailMessage {
  id: string;
  storeId?: string;
  subject: string;
  sender: string;
  recipients: string[];
  body: string;
  receivedTime: Date;
  isRead: boolean;
  hasAttachments: boolean;
}

export interface EmailAttachment {
  /** Filename shown in Outlook (e.g. "meme.jpg"). Required with content_base64. */
  name?: string;
  /** Base64-encoded file content; written to a temp file and attached by value. */
  content_base64?: string;
  /** Absolute Windows path to an existing file — alternative to content_base64. */
  path?: string;
  /** Content-ID: makes this an inline attachment referenced as <img src="cid:VALUE"> in the HTML body. */
  cid?: string;
}

export interface EmailDraft {
  to: string[];
  cc?: string[];
  bcc?: string[];
  subject: string;
  body: string;
  isHtml?: boolean;
  attachments?: EmailAttachment[];
}

/**
 * Render a date as a naive local timestamp for PowerShell.
 *
 * `toISOString()` cannot be used here: PowerShell parses the trailing `Z` and
 * converts to local time, so a window starting at UTC midnight begins at 01:00
 * in a UTC+1 zone and events in that first hour fall outside the filter.
 */
function toPowerShellDate(date: Date): string {
  const pad = (value: number) => String(value).padStart(2, '0');
  return `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}` +
    `T${pad(date.getHours())}:${pad(date.getMinutes())}:${pad(date.getSeconds())}`;
}

export class OutlookManager {
  private powershellPath: string;

  constructor() {
    this.powershellPath = 'powershell.exe';
  }

  private async executePowerShell(script: string): Promise<string> {
    return new Promise((resolve, reject) => {
      // Prepare UTF-8 encoded script
      const utf8Script = `
        chcp 65001 > $null
        [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
        [Console]::InputEncoding = [System.Text.Encoding]::UTF8
        $OutputEncoding = [System.Text.Encoding]::UTF8
        ${script}
      `;

      const ps = spawn(this.powershellPath, [
        '-NoProfile',
        '-NonInteractive',
        '-ExecutionPolicy', 'Bypass',
        '-Command', utf8Script
      ], {
        env: { 
          ...process.env, 
          'PYTHONIOENCODING': 'utf-8'
        }
      });

      let stdout = '';
      let stderr = '';

      ps.stdout.setEncoding('utf8');
      ps.stderr.setEncoding('utf8');

      ps.stdout.on('data', (data) => {
        stdout += data;
      });

      ps.stderr.on('data', (data) => {
        stderr += data;
      });

      ps.on('close', (code) => {
        if (code === 0) {
          // Clean the output
          let cleanOutput = stdout
            .replace(/^\uFEFF/, '') // Remove BOM
            .replace(/[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]/g, '') // Remove control characters
            .trim();
          
          resolve(cleanOutput);
        } else {
          reject(new Error(`PowerShell failed (code ${code}): ${stderr}`));
        }
      });
    });
  }

  /**
   * Common email retrieval function
   */
  private async getEmailsFromFolder(folderType: number, count: number = 10, sortBy: string = "[ReceivedTime]"): Promise<EmailMessage[]> {
    const script = `
      try {
        Add-Type -AssemblyName "Microsoft.Office.Interop.Outlook" -ErrorAction Stop
        $outlook = New-Object -ComObject Outlook.Application -ErrorAction Stop
        $namespace = $outlook.GetNamespace("MAPI")
        $folder = $namespace.GetDefaultFolder(${folderType})
        
        if ($folder.Items.Count -eq 0) {
          Write-Output "[]"
          exit 0
        }
        
        $items = $folder.Items
        $items.Sort("${sortBy}", $true)
        
        $emails = @()
        $counter = 0
        
        foreach ($item in $items) {
          if ($counter -ge ${count}) { break }
          
          try {
            $subject = if ($item.Subject) { $item.Subject.ToString() -replace '[\\x00-\\x1F\\x7F]', '' } else { "No Subject" }
            $sender = if ($item.SenderEmailAddress) { $item.SenderEmailAddress.ToString() -replace '[\\x00-\\x1F\\x7F]', '' } else { "Unknown" }
            $body = if ($item.Body) { 
              $bodyStr = $item.Body.ToString() -replace '[\\x00-\\x1F\\x7F]', ''
              if ($bodyStr.Length -gt 150) { $bodyStr.Substring(0, 150) + "..." } else { $bodyStr }
            } else { "" }
            
            $timeStamp = if ($item.SentOn -and ${folderType} -eq 5) { 
              $item.SentOn.ToString("yyyy-MM-dd HH:mm:ss") 
            } elseif ($item.ReceivedTime) { 
              $item.ReceivedTime.ToString("yyyy-MM-dd HH:mm:ss") 
            } else { 
              (Get-Date).ToString("yyyy-MM-dd HH:mm:ss") 
            }
            
            $emails += [PSCustomObject]@{
              Id = if ($item.EntryID) { $item.EntryID.ToString() } else { "no-id-$counter" }
              StoreID = if ($item.Session -and $item.Session.DefaultStore -and $item.Session.DefaultStore.StoreID) { 
                $item.Session.DefaultStore.StoreID.ToString() 
              } elseif ($item.Parent -and $item.Parent.StoreID) { 
                $item.Parent.StoreID.ToString() 
              } else { 
                try { $namespace.DefaultStore.StoreID.ToString() } catch { "" }
              }
              Subject = $subject
              Sender = $sender
              Recipients = @()
              Body = $body
              ReceivedTime = $timeStamp
              IsRead = if (${folderType} -eq 5) { $true } else { -not $item.UnRead }
              HasAttachments = $item.Attachments.Count -gt 0
            }
            
            $counter++
          } catch { $counter++; continue }
        }
        
        if ($emails.Count -eq 0) { Write-Output "[]" } 
        else { Write-Output ($emails | ConvertTo-Json -Depth 2 -Compress) }
        
      } catch {
        Write-Output ([PSCustomObject]@{ error = $_.Exception.Message; type = "OutlookConnectionError" } | ConvertTo-Json -Compress)
      }
    `;

    try {
      const result = await this.executePowerShell(script);
      if (!result || result.trim() === '') return [];
      
      const cleanResult = result.replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\u007F]/g, '').trim();
      const parsed = JSON.parse(cleanResult);
      
      if (parsed.error) throw new Error(`Outlook Error: ${parsed.error}`);
      
      const emailArray = Array.isArray(parsed) ? parsed : [parsed];
      return emailArray.map((item: any) => ({
        id: this.cleanText(item.Id || ''),
        storeId: this.cleanText(item.StoreID || ''),
        subject: this.cleanText(item.Subject || 'No Subject'),
        sender: this.cleanText(item.Sender || 'Unknown'),
        recipients: [],
        body: this.cleanText(item.Body || ''),
        receivedTime: new Date(item.ReceivedTime),
        isRead: Boolean(item.IsRead),
        hasAttachments: Boolean(item.HasAttachments)
      }));
      
    } catch (error) {
      console.error('Email fetch failed:', error);
      return [{
        id: 'fallback-1',
        storeId: '',
        subject: 'Email content unavailable',
        sender: 'system@outlook.com',
        recipients: [],
        body: 'Unable to retrieve email content.',
        receivedTime: new Date(),
        isRead: true,
        hasAttachments: false
      }];
    }
  }

  private cleanText(text: string): string {
    if (!text) return '';
    
    return text
      .replace(/[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]/g, '') // Remove control chars
      .replace(/\r\n/g, ' ') // Replace CRLF with space
      .replace(/[\r\n]/g, ' ') // Replace any remaining line breaks
      .replace(/\s+/g, ' ') // Collapse multiple spaces
      .trim();
  }

  private formatBodyForOutlook(body: string): string {
    if (!body) return '';
    
    // Normalize line breaks to Windows format (CRLF) for Outlook
    return body
      .replace(/[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]/g, '') // Remove control chars except CR and LF
      .replace(/\r\n/g, '\n') // Normalize to LF first
      .replace(/\r/g, '\n') // Convert any remaining CR to LF
      .replace(/\n/g, '\r\n') // Convert all LF to CRLF for Windows
      .trim();
  }

  async getInboxEmails(count: number = 10): Promise<EmailMessage[]> {
    return this.getEmailsFromFolder(6, count, "[ReceivedTime]"); // 6 = Inbox
  }

  async getSentEmails(count: number = 10): Promise<EmailMessage[]> {
    return this.getEmailsFromFolder(5, count, "[SentOn]"); // 5 = Sent Items
  }

  async getDraftEmails(count: number = 10): Promise<EmailMessage[]> {
    return this.getEmailsFromFolder(16, count, "[LastModificationTime]"); // 16 = Drafts
  }

  async getEmailById(id: string): Promise<EmailMessage> {
    const script = `
      try {
        Add-Type -AssemblyName "Microsoft.Office.Interop.Outlook" -ErrorAction Stop
        $outlook = New-Object -ComObject Outlook.Application -ErrorAction Stop
        $namespace = $outlook.GetNamespace("MAPI")
        
        # Try GetItemFromID first (fastest method)
        $item = $null
        try {
          $item = $namespace.GetItemFromID("${id.replace(/"/g, '""')}")
        } catch {
          # Fallback: search through folders
          foreach ($folderNum in @(6, 5, 16)) {
            $folder = $namespace.GetDefaultFolder($folderNum)
            foreach ($email in $folder.Items) {
              if ($email.EntryID -eq "${id.replace(/"/g, '""')}") {
                $item = $email
                break
              }
            }
            if ($item) { break }
          }
        }
        
        if (-not $item) { throw "Email not found" }
        
        # Extract data
        $subject = if ($item.Subject) { $item.Subject } else { "No Subject" }
        $sender = if ($item.SenderEmailAddress) { $item.SenderEmailAddress } else { "Unknown" }
        $body = if ($item.Body) { $item.Body } else { "" }
        $recipients = @()
        if ($item.Recipients) {
          foreach ($r in $item.Recipients) {
            $addr = if ($r.Address) { $r.Address } else { $r.Name }
            if ($addr) { $recipients += $addr }
          }
        }
        $timestamp = if ($item.SentOn) { $item.SentOn.ToString("yyyy-MM-dd HH:mm:ss") } 
                     elseif ($item.ReceivedTime) { $item.ReceivedTime.ToString("yyyy-MM-dd HH:mm:ss") } 
                     else { (Get-Date).ToString("yyyy-MM-dd HH:mm:ss") }
        
        Write-Output ([PSCustomObject]@{
          Id = "${id.replace(/"/g, '""')}"
          Subject = $subject
          Sender = $sender
          Recipients = $recipients
          Body = $body
          ReceivedTime = $timestamp
          IsRead = -not $item.UnRead
          HasAttachments = $item.Attachments.Count -gt 0
          Success = $true
        } | ConvertTo-Json -Depth 3 -Compress)
        
      } catch {
        Write-Output ([PSCustomObject]@{
          Id = "${id.replace(/"/g, '""')}"
          Subject = "Email not found"
          Sender = "system"
          Recipients = @()
          Body = "Error: $($_.Exception.Message)"
          ReceivedTime = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
          IsRead = $true
          HasAttachments = $false
          Success = $false
        } | ConvertTo-Json -Depth 3 -Compress)
      }
    `;

    try {
      const result = await this.executePowerShell(script);
      const cleanResult = result.replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\u007F]/g, '').trim();
      const emailData = JSON.parse(cleanResult);
      
      return {
        id: emailData.Id || id,
        subject: this.cleanText(emailData.Subject || 'No Subject'),
        sender: this.cleanText(emailData.Sender || 'Unknown Sender'),
        recipients: Array.isArray(emailData.Recipients) ? emailData.Recipients.map((r: any) => this.cleanText(r)) : [],
        body: emailData.Body || '',
        receivedTime: new Date(emailData.ReceivedTime || new Date()),
        isRead: Boolean(emailData.IsRead),
        hasAttachments: Boolean(emailData.HasAttachments)
      };
      
    } catch (error) {
      return {
        id: id,
        subject: 'Email parsing failed',
        sender: 'system',
        recipients: [],
        body: `Failed to parse email: ${error instanceof Error ? error.message : String(error)}`,
        receivedTime: new Date(),
        isRead: true,
        hasAttachments: false
      };
    }
  }

  async createDraft(draft: EmailDraft): Promise<string> {
    const cleanSubject = this.cleanText(draft.subject);
    // For HTML: write body to a temp file so PowerShell reads it with proper UTF-8
    // encoding, avoiding all command-line length limits and quoting issues.
    // For plain text: use the normal CRLF-formatted inline string.
    let bodyAssignment: string;
    if (draft.isHtml) {
      const tmpFilePath = join(tmpdir(), `fhacc-draft-${Date.now()}.html`);
      writeFileSync(tmpFilePath, draft.body, 'utf8');
      const winPath = tmpFilePath.replace(/\\/g, '\\\\');
      bodyAssignment = `$mail.BodyFormat = 2
        $mail.HTMLBody = [System.IO.File]::ReadAllText("${winPath}", [System.Text.Encoding]::UTF8)
        Remove-Item "${winPath}" -ErrorAction SilentlyContinue`;
    } else {
      const formattedBody = this.formatBodyForOutlook(draft.body);
      bodyAssignment = `$mail.Body = "${formattedBody.replace(/"/g, '""')}"`;
    }

    // Attachments: base64 payloads land in a per-draft temp dir (so Outlook picks
    // up the real filename as the display name); paths are attached as-is.
    // olByValue copies the file into the item at Add() time, so temp files can be
    // deleted as soon as the script finishes. A cid turns the attachment into an
    // inline image (PR_ATTACH_CONTENT_ID) hidden from the paperclip list
    // (PR_ATTACHMENT_HIDDEN) — attach BEFORE setting HTMLBody so cid: references resolve.
    const tempDirs: string[] = [];
    const attachmentLines: string[] = [];
    for (const [i, att] of (draft.attachments ?? []).entries()) {
      let filePath: string;
      if (att.content_base64) {
        const safeName = (att.name || `attachment-${i + 1}`).replace(/[\\/:*?"<>|]/g, '_');
        const dir = mkdtempSync(join(tmpdir(), 'fhacc-att-'));
        tempDirs.push(dir);
        filePath = join(dir, safeName);
        writeFileSync(filePath, Buffer.from(att.content_base64, 'base64'));
      } else if (att.path) {
        filePath = att.path;
      } else {
        tempDirs.forEach(d => rmSync(d, { recursive: true, force: true }));
        throw new Error(`Attachment ${i + 1}: either content_base64 or path is required`);
      }
      const winPath = filePath.replace(/\\/g, '\\\\');
      attachmentLines.push(`$att = $mail.Attachments.Add("${winPath}", 1)`);
      if (att.cid) {
        const safeCid = att.cid.replace(/[^A-Za-z0-9._@-]/g, '');
        attachmentLines.push(
          `$att.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "${safeCid}")`,
          `try { $att.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x7FFE000B", $true) } catch {}`
        );
      }
    }

    const recipientBlock = (addresses: string[] | undefined, olType: number) => {
      const list = (addresses ?? []).filter(a => a && a.trim());
      if (list.length === 0) return '';
      return `
        foreach ($recipient in @("${list.join('","')}")) {
          $r = $mail.Recipients.Add($recipient.Trim())
          $r.Type = ${olType}
        }`;
    };

    const script = `
      try {
        Add-Type -AssemblyName "Microsoft.Office.Interop.Outlook"
        $outlook = New-Object -ComObject Outlook.Application
        $mail = $outlook.CreateItem(0)

        $mail.Subject = "${cleanSubject.replace(/"/g, '""')}"
        ${attachmentLines.join('\n        ')}
        ${bodyAssignment}
        ${recipientBlock(draft.to, 1)}${recipientBlock(draft.cc, 2)}${recipientBlock(draft.bcc, 3)}

        $mail.Recipients.ResolveAll() | Out-Null
        $mail.Save()

        Write-Output "success"
      } catch {
        Write-Output "error: $($_.Exception.Message)"
      }
    `;

    try {
      const result = await this.executePowerShell(script);
      if (result.startsWith('error:')) {
        throw new Error(result.substring(7));
      }
      return 'Draft created successfully';
    } finally {
      tempDirs.forEach(d => rmSync(d, { recursive: true, force: true }));
    }
  }

  async markAsRead(id: string): Promise<void> {
    return Promise.resolve();
  }

  async searchInboxEmails(query: string, count: number = 10): Promise<EmailMessage[]> {
    const emails = await this.getInboxEmails(Math.min(count * 2, 50));
    const { EmailSummarizer } = await import('./email-summarizer.js');
    const searchResults = EmailSummarizer.searchEmails(emails, query);
    return searchResults.slice(0, count);
  }

  async searchSentEmails(query: string, count: number = 10): Promise<EmailMessage[]> {
    const emails = await this.getSentEmails(Math.min(count * 2, 50));
    const { EmailSummarizer } = await import('./email-summarizer.js');
    const searchResults = EmailSummarizer.searchEmails(emails, query);
    return searchResults.slice(0, count);
  }

  async searchDraftEmails(query: string, count: number = 10): Promise<EmailMessage[]> {
    const emails = await this.getDraftEmails(Math.min(count * 2, 50));
    const { EmailSummarizer } = await import('./email-summarizer.js');
    const searchResults = EmailSummarizer.searchEmails(emails, query);
    return searchResults.slice(0, count);
  }

  /**
   * Duplicate an existing email to create a new draft
   * Uses ReplyAll method to preserve complete formatting
   */
  async duplicateEmailAsDraft(sourceEmailId: string, newSubject?: string, newRecipients?: string[], storeId?: string): Promise<string> {
    const script = `
      try {
        Add-Type -AssemblyName "Microsoft.Office.Interop.Outlook" -ErrorAction Stop
        $outlook = New-Object -ComObject Outlook.Application -ErrorAction Stop
        $namespace = $outlook.GetNamespace("MAPI")
        
        # Find the original email using EntryID and StoreID
        $item = $null
        $sourceEntryID = "${sourceEmailId.replace(/"/g, '""')}"
        $sourceStoreID = "${(storeId || '').replace(/"/g, '""')}"
        
        # Try GetItemFromID with StoreID first, then fallback methods
        try {
          if ($sourceStoreID -and $sourceStoreID.Length -gt 0) {
            $item = $namespace.GetItemFromID($sourceEntryID, $sourceStoreID)
          } else {
            $item = $namespace.GetItemFromID($sourceEntryID)
          }
        } catch {
          # Fallback: search through folders
          $folders = @(
            $namespace.GetDefaultFolder(6),  # Inbox
            $namespace.GetDefaultFolder(5),  # Sent Items
            $namespace.GetDefaultFolder(16)  # Drafts
          )
          
          foreach ($folder in $folders) {
            try {
              foreach ($email in $folder.Items) {
                if ($email.EntryID -eq $sourceEntryID) {
                  $item = $email
                  break
                }
              }
              if ($item) { break }
            } catch { continue }
          }
        }
        
        if (-not $item) {
          throw "Original email not found with EntryID: $sourceEntryID"
        }
        
        # Use ReplyAll to preserve all formatting, then modify
        $draft = $item.ReplyAll()
        
        # Update subject if provided
        $subjectToUse = "${(newSubject || '').replace(/"/g, '""')}"
        if ($subjectToUse.Length -gt 0) {
          $draft.Subject = $subjectToUse
        }
        
        # Clear and set recipients if provided
        if ("${(newRecipients || []).join(',')}" -ne "") {
          $draft.Recipients.RemoveAll()
          $newRecipientsList = @("${(newRecipients || []).join('","')}")
          foreach ($recipient in $newRecipientsList) {
            if ($recipient.Trim().Length -gt 0) {
              $draft.Recipients.Add($recipient.Trim()) | Out-Null
            }
          }
        }
        
        # Resolve recipients and save
        try {
          $draft.Recipients.ResolveAll() | Out-Null
        } catch { }
        
        $draft.Save()
        Write-Output "success"
        
      } catch {
        Write-Output "error: $($_.Exception.Message)"
      }
    `;
    
    const result = await this.executePowerShell(script);
    if (result.startsWith('error:')) {
      throw new Error(result.substring(7));
    }
    return 'Draft created successfully using ReplyAll method';
  }

  /**
   * Set Show As (Free/Busy status) for calendar events
   */
  async setShowAs(options: {
    eventId?: string;
    subject?: string;
    startDate?: Date;
    showAs: 'Free' | 'Tentative' | 'Busy' | 'OutOfOffice' | 'WorkingElsewhere';
  }): Promise<{ success: boolean; message: string }> {
    try {
      // Map ShowAs values to Outlook constants
      const showAsMap: Record<string, number> = {
        'Free': 0,           // olFree
        'Tentative': 1,      // olTentative
        'Busy': 2,           // olBusy
        'OutOfOffice': 3,    // olOutOfOffice
        'WorkingElsewhere': 4 // olWorkingElsewhere
      };

      const busyStatus = showAsMap[options.showAs];
      const eventId = options.eventId ? `"${options.eventId.replace(/"/g, '""')}"` : 'null';
      const subject = options.subject ? `"${options.subject.replace(/"/g, '""')}"` : 'null';
      const startDate = options.startDate ? `"${options.startDate.toISOString()}"` : 'null';

      const script = `
        try {
          Add-Type -AssemblyName "Microsoft.Office.Interop.Outlook" -ErrorAction Stop
          $outlook = New-Object -ComObject Outlook.Application -ErrorAction Stop
          $namespace = $outlook.GetNamespace("MAPI")
          $calendar = $namespace.GetDefaultFolder(9)
          
          $appointmentItem = $null
          
          # Find by ID if provided
          if (${eventId} -ne $null) {
            try {
              $appointmentItem = $namespace.GetItemFromID(${eventId})
            } catch { }
          }
          
          # Search by subject or date if not found
          if (-not $appointmentItem) {
            $items = $calendar.Items
            $items.Sort("[Start]")
            
            foreach ($item in $items) {
              $matchSubject = (${subject} -eq $null) -or ($item.Subject -like "*$(${subject})*")
              $matchDate = (${startDate} -eq $null) -or ([Math]::Abs(([DateTime]$item.Start - [DateTime]${startDate}).TotalMinutes) -lt 1)
              
              if ($matchSubject -and $matchDate) {
                $appointmentItem = $item
                break
              }
            }
          }
          
          if (-not $appointmentItem) {
            throw "Appointment not found. Please provide eventId, subject, or startDate."
          }
          
          $appointmentItem.BusyStatus = ${busyStatus}
          $appointmentItem.Save()
          
          Write-Output ([PSCustomObject]@{
            Success = $true
            Subject = $appointmentItem.Subject
            ShowAs = "${options.showAs}"
          } | ConvertTo-Json -Compress)
          
        } catch {
          Write-Output ([PSCustomObject]@{
            Success = $false
            Error = $_.Exception.Message
          } | ConvertTo-Json -Compress)
        }
      `;

      const result = await this.executePowerShell(script);
      const cleanResult = result.replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\u007F]/g, '').trim();
      const data = JSON.parse(cleanResult);

      if (!data.Success) {
        throw new Error(data.Error || 'Failed to set Show As');
      }

      return {
        success: true,
        message: `Show As set to ${options.showAs} for: ${data.Subject}`
      };
    } catch (error) {
      throw new Error(`Failed to set Show As: ${error instanceof Error ? error.message : String(error)}`);
    }
  }

  /**
   * Create calendar event with Show As setting
   */
  async createEventWithShowAs(options: {
    subject: string;
    start: Date;
    end: Date;
    location?: string;
    body?: string;
    showAs?: 'Free' | 'Tentative' | 'Busy' | 'OutOfOffice' | 'WorkingElsewhere';
    reminderMinutes?: number;
  }): Promise<{ success: boolean; eventId: string; message: string }> {
    try {
      // Map ShowAs values to Outlook constants
      const showAsMap: Record<string, number> = {
        'Free': 0,
        'Tentative': 1,
        'Busy': 2,
        'OutOfOffice': 3,
        'WorkingElsewhere': 4
      };

      const busyStatus = showAsMap[options.showAs || 'Busy'];
      const cleanSubject = this.cleanText(options.subject);
      const cleanLocation = options.location ? this.cleanText(options.location) : '';
      const cleanBody = options.body ? this.cleanText(options.body) : '';
      
      const script = `
        try {
          Add-Type -AssemblyName "Microsoft.Office.Interop.Outlook" -ErrorAction Stop
          $outlook = New-Object -ComObject Outlook.Application -ErrorAction Stop
          $appointmentItem = $outlook.CreateItem(1)
          
          $appointmentItem.Subject = "${cleanSubject.replace(/"/g, '""')}"
          $appointmentItem.Start = [DateTime]"${options.start.toISOString()}"
          $appointmentItem.End = [DateTime]"${options.end.toISOString()}"
          
          ${options.location ? `$appointmentItem.Location = "${cleanLocation.replace(/"/g, '""')}"` : ''}
          ${options.body ? `$appointmentItem.Body = "${cleanBody.replace(/"/g, '""')}"` : ''}
          
          $appointmentItem.BusyStatus = ${busyStatus}
          
          ${options.reminderMinutes !== undefined ? `
          $appointmentItem.ReminderSet = $true
          $appointmentItem.ReminderMinutesBeforeStart = ${options.reminderMinutes}
          ` : ''}
          
          $appointmentItem.Save()
          
          Write-Output ([PSCustomObject]@{
            Success = $true
            EventId = $appointmentItem.EntryID
            Subject = $appointmentItem.Subject
            ShowAs = "${options.showAs || 'Busy'}"
          } | ConvertTo-Json -Compress)
          
        } catch {
          Write-Output ([PSCustomObject]@{
            Success = $false
            Error = $_.Exception.Message
          } | ConvertTo-Json -Compress)
        }
      `;

      const result = await this.executePowerShell(script);
      const cleanResult = result.replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\u007F]/g, '').trim();
      const data = JSON.parse(cleanResult);

      if (!data.Success) {
        throw new Error(data.Error || 'Failed to create event');
      }

      return {
        success: true,
        eventId: data.EventId,
        message: `Event created: ${data.Subject} with Show As: ${data.ShowAs}`
      };
    } catch (error) {
      throw new Error(`Failed to create event with Show As: ${error instanceof Error ? error.message : String(error)}`);
    }
  }

  /**
   * List calendar events within a date range
   */
  async listEvents(options: {
    startDate: Date;
    endDate?: Date;
    calendar?: string;
  }): Promise<any[]> {
    try {
      const endDate = options.endDate || options.startDate;
      const calendarName = options.calendar || '';
      
      const script = `
        try {
          Add-Type -AssemblyName "Microsoft.Office.Interop.Outlook" -ErrorAction Stop
          $outlook = New-Object -ComObject Outlook.Application -ErrorAction Stop
          $namespace = $outlook.GetNamespace("MAPI")
          
          # Get calendar
          ${calendarName ? `
          $calendar = $null
          foreach ($folder in $namespace.Folders) {
            if ($folder.Name -eq "${calendarName.replace(/"/g, '""')}") {
              $calendar = $folder.Folders("Calendar")
              break
            }
          }
          if (-not $calendar) { throw "Calendar not found: ${calendarName.replace(/"/g, '""')}" }
          ` : `
          $calendar = $namespace.GetDefaultFolder(9)
          `}
          
          # Create filter for date range. Match anything overlapping the window,
          # not only appointments contained by it.
          $startDate = [DateTime]"${toPowerShellDate(options.startDate)}"
          $endDate = ([DateTime]"${toPowerShellDate(endDate)}").AddDays(1)
          $filter = "[Start] < '$($endDate.ToString('g'))' AND [End] > '$($startDate.ToString('g'))'"

          # Get events. Sorting by [Start] and opting into recurrences must both
          # happen before Restrict, otherwise Outlook returns only the master
          # appointment of each series and every occurrence is invisible.
          $items = $calendar.Items
          $items.Sort("[Start]")
          $items.IncludeRecurrences = $true
          $items = $items.Restrict($filter)

          # Build JSON array
          $events = @()
          foreach ($item in $items) {
            $events += [PSCustomObject]@{
              Id = $item.EntryID
              Subject = $item.Subject
              Start = $item.Start.ToString("yyyy-MM-ddTHH:mm:ss")
              End = $item.End.ToString("yyyy-MM-ddTHH:mm:ss")
              Location = if ($item.Location) { $item.Location } else { "" }
              Body = if ($item.Body) { $item.Body } else { "" }
              BusyStatus = $item.BusyStatus
              IsAllDayEvent = $item.AllDayEvent
              Organizer = if ($item.Organizer) { $item.Organizer } else { "" }
              RequiredAttendees = if ($item.RequiredAttendees) { $item.RequiredAttendees } else { "" }
            }
          }
          
          Write-Output ($events | ConvertTo-Json -Compress)
          
        } catch {
          Write-Output ([PSCustomObject]@{
            Error = $_.Exception.Message
          } | ConvertTo-Json -Compress)
        }
      `;

      const result = await this.executePowerShell(script);
      const cleanResult = result.replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\u007F]/g, '').trim();
      
      if (!cleanResult || cleanResult === '') {
        return [];
      }
      
      const data = JSON.parse(cleanResult);
      
      if (data.Error) {
        throw new Error(data.Error);
      }
      
      return Array.isArray(data) ? data : [data];
    } catch (error) {
      throw new Error(`Failed to list events: ${error instanceof Error ? error.message : String(error)}`);
    }
  }

  /**
   * Update an existing calendar event
   */
  async updateEvent(options: {
    eventId: string;
    subject?: string;
    startDate?: string;
    startTime?: string;
    endDate?: string;
    endTime?: string;
    location?: string;
    body?: string;
    calendar?: string;
  }): Promise<{ success: boolean; message: string }> {
    try {
      const calendarName = options.calendar || '';
      let startDateTime = '';
      let endDateTime = '';
      
      if (options.startDate && options.startTime) {
        startDateTime = `${options.startDate} ${options.startTime}`;
      }
      
      if (options.endDate && options.endTime) {
        endDateTime = `${options.endDate} ${options.endTime}`;
      } else if (startDateTime) {
        // Default to 30 minutes after start
        endDateTime = 'ADD30MIN';
      }
      
      const script = `
        try {
          Add-Type -AssemblyName "Microsoft.Office.Interop.Outlook" -ErrorAction Stop
          $outlook = New-Object -ComObject Outlook.Application -ErrorAction Stop
          $namespace = $outlook.GetNamespace("MAPI")
          
          # Get appointment
          $appointment = $namespace.GetItemFromID("${options.eventId.replace(/"/g, '""')}")
          
          if (-not $appointment) {
            throw "Event not found with ID: ${options.eventId.replace(/"/g, '""')}"
          }
          
          # Update properties
          ${options.subject ? `$appointment.Subject = "${options.subject.replace(/"/g, '""')}"` : ''}
          ${startDateTime ? `$appointment.Start = [DateTime]"${startDateTime.replace(/"/g, '""')}"` : ''}
          ${endDateTime === 'ADD30MIN' ? `$appointment.End = $appointment.Start.AddMinutes(30)` : endDateTime ? `$appointment.End = [DateTime]"${endDateTime.replace(/"/g, '""')}"` : ''}
          ${options.location ? `$appointment.Location = "${options.location.replace(/"/g, '""')}"` : ''}
          ${options.body ? `$appointment.Body = "${options.body.replace(/"/g, '""')}"` : ''}
          
          $appointment.Save()
          
          Write-Output ([PSCustomObject]@{
            Success = $true
          } | ConvertTo-Json -Compress)
          
        } catch {
          Write-Output ([PSCustomObject]@{
            Success = $false
            Error = $_.Exception.Message
          } | ConvertTo-Json -Compress)
        }
      `;

      const result = await this.executePowerShell(script);
      const cleanResult = result.replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\u007F]/g, '').trim();
      const data = JSON.parse(cleanResult);

      if (!data.Success) {
        throw new Error(data.Error || 'Failed to update event');
      }

      return {
        success: true,
        message: 'Event updated successfully'
      };
    } catch (error) {
      throw new Error(`Failed to update event: ${error instanceof Error ? error.message : String(error)}`);
    }
  }

  /**
   * Delete a calendar event
   */
  async deleteEvent(options: {
    eventId: string;
    calendar?: string;
  }): Promise<{ success: boolean; message: string }> {
    try {
      const script = `
        try {
          Add-Type -AssemblyName "Microsoft.Office.Interop.Outlook" -ErrorAction Stop
          $outlook = New-Object -ComObject Outlook.Application -ErrorAction Stop
          $namespace = $outlook.GetNamespace("MAPI")
          
          # Get appointment
          $appointment = $namespace.GetItemFromID("${options.eventId.replace(/"/g, '""')}")
          
          if (-not $appointment) {
            throw "Event not found with ID: ${options.eventId.replace(/"/g, '""')}"
          }
          
          # Delete the appointment
          $appointment.Delete()
          
          Write-Output ([PSCustomObject]@{
            Success = $true
          } | ConvertTo-Json -Compress)
          
        } catch {
          Write-Output ([PSCustomObject]@{
            Success = $false
            Error = $_.Exception.Message
          } | ConvertTo-Json -Compress)
        }
      `;

      const result = await this.executePowerShell(script);
      const cleanResult = result.replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\u007F]/g, '').trim();
      const data = JSON.parse(cleanResult);

      if (!data.Success) {
        throw new Error(data.Error || 'Failed to delete event');
      }

      return {
        success: true,
        message: 'Event deleted successfully'
      };
    } catch (error) {
      throw new Error(`Failed to delete event: ${error instanceof Error ? error.message : String(error)}`);
    }
  }

  /**
   * Find free time slots in calendar
   */
  async findFreeSlots(options: {
    startDate: Date;
    endDate?: Date;
    duration?: number;
    workDayStart?: number;
    workDayEnd?: number;
    calendar?: string;
  }): Promise<any[]> {
    try {
      const endDate = options.endDate || new Date(options.startDate.getTime() + 7 * 24 * 60 * 60 * 1000);
      const duration = options.duration || 30;
      const workDayStart = options.workDayStart || 9;
      const workDayEnd = options.workDayEnd || 17;
      const calendarName = options.calendar || '';
      
      const script = `
        try {
          Add-Type -AssemblyName "Microsoft.Office.Interop.Outlook" -ErrorAction Stop
          $outlook = New-Object -ComObject Outlook.Application -ErrorAction Stop
          $namespace = $outlook.GetNamespace("MAPI")
          
          # Get calendar
          ${calendarName ? `
          $calendar = $null
          foreach ($folder in $namespace.Folders) {
            if ($folder.Name -eq "${calendarName.replace(/"/g, '""')}") {
              $calendar = $folder.Folders("Calendar")
              break
            }
          }
          if (-not $calendar) { throw "Calendar not found: ${calendarName.replace(/"/g, '""')}" }
          ` : `
          $calendar = $namespace.GetDefaultFolder(9)
          `}
          
          # Get all events in date range
          $startDate = [DateTime]"${toPowerShellDate(options.startDate)}"
          $endDate = [DateTime]"${toPowerShellDate(endDate)}"
          $filterEnd = $endDate.AddDays(1)
          $filter = "[Start] < '$($filterEnd.ToString('g'))' AND [End] > '$($startDate.ToString('g'))'"

          # Sorting by [Start] and opting into recurrences must both happen
          # before Restrict, otherwise recurring meetings never show up as busy
          # and the whole day is reported free.
          $items = $calendar.Items
          $items.Sort("[Start]")
          $items.IncludeRecurrences = $true
          $items = $items.Restrict($filter)

          # Build busy slots array
          $busySlots = @()
          foreach ($item in $items) {
            if ($item.BusyStatus -eq 1 -or $item.BusyStatus -eq 2 -or $item.BusyStatus -eq 3) { # Tentative, Busy or OutOfOffice
              $busySlots += [PSCustomObject]@{
                Start = $item.Start
                End = $item.End
              }
            }
          }
          
          # Find free slots
          $freeSlots = @()
          $currentDate = $startDate.Date
          
          while ($currentDate -le $endDate.Date) {
            # Skip weekends
            $dayOfWeek = $currentDate.DayOfWeek
            if ($dayOfWeek -ne [DayOfWeek]::Saturday -and $dayOfWeek -ne [DayOfWeek]::Sunday) {
              $slotStart = $currentDate.AddHours(${workDayStart})
              $workDayEndTime = $currentDate.AddHours(${workDayEnd})
              
              while ($slotStart.AddMinutes(${duration}) -le $workDayEndTime) {
                $slotEnd = $slotStart.AddMinutes(${duration})
                
                # Check if slot is free
                $isFree = $true
                foreach ($busy in $busySlots) {
                  if ($slotStart -lt $busy.End -and $slotEnd -gt $busy.Start) {
                    $isFree = $false
                    break
                  }
                }
                
                if ($isFree) {
                  $freeSlots += [PSCustomObject]@{
                    Start = $slotStart.ToString("yyyy-MM-ddTHH:mm:ss")
                    End = $slotEnd.ToString("yyyy-MM-ddTHH:mm:ss")
                  }
                }
                
                $slotStart = $slotStart.AddMinutes(30)
              }
            }
            
            $currentDate = $currentDate.AddDays(1)
          }
          
          Write-Output ($freeSlots | ConvertTo-Json -Compress)
          
        } catch {
          Write-Output ([PSCustomObject]@{
            Error = $_.Exception.Message
          } | ConvertTo-Json -Compress)
        }
      `;

      const result = await this.executePowerShell(script);
      const cleanResult = result.replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\u007F]/g, '').trim();
      
      if (!cleanResult || cleanResult === '' || cleanResult === '[]') {
        return [];
      }
      
      const data = JSON.parse(cleanResult);
      
      if (data.Error) {
        throw new Error(data.Error);
      }
      
      return Array.isArray(data) ? data : [data];
    } catch (error) {
      throw new Error(`Failed to find free slots: ${error instanceof Error ? error.message : String(error)}`);
    }
  }

  /**
   * Get attendee status for a meeting
   */
  async getAttendeeStatus(options: {
    eventId: string;
    calendar?: string;
  }): Promise<any> {
    try {
      const script = `
        try {
          Add-Type -AssemblyName "Microsoft.Office.Interop.Outlook" -ErrorAction Stop
          $outlook = New-Object -ComObject Outlook.Application -ErrorAction Stop
          $namespace = $outlook.GetNamespace("MAPI")
          
          # Get appointment
          $appointment = $namespace.GetItemFromID("${options.eventId.replace(/"/g, '""')}")
          
          if (-not $appointment) {
            throw "Event not found with ID: ${options.eventId.replace(/"/g, '""')}"
          }
          
          # Check if it's a meeting
          if ($appointment.MeetingStatus -eq 0) { # olNonMeeting
            throw "The specified event is not a meeting"
          }
          
          # Get attendees
          $attendees = @()
          foreach ($recipient in $appointment.Recipients) {
            $responseStatus = switch ($recipient.MeetingResponseStatus) {
              1 { "Organizer" }
              2 { "Tentative" }
              3 { "Accepted" }
              4 { "Declined" }
              0 { "Not Responded" }
              default { "Unknown" }
            }
            
            $attendees += [PSCustomObject]@{
              Name = $recipient.Name
              Email = if ($recipient.Address) { $recipient.Address } else { $recipient.Name }
              ResponseStatus = $responseStatus
            }
          }
          
          # Build result
          $result = [PSCustomObject]@{
            Subject = $appointment.Subject
            Start = $appointment.Start.ToString("yyyy-MM-ddTHH:mm:ss")
            End = $appointment.End.ToString("yyyy-MM-ddTHH:mm:ss")
            Location = if ($appointment.Location) { $appointment.Location } else { "" }
            Organizer = if ($appointment.Organizer) { $appointment.Organizer } else { "" }
            Attendees = $attendees
          }
          
          Write-Output ($result | ConvertTo-Json -Compress -Depth 3)
          
        } catch {
          Write-Output ([PSCustomObject]@{
            Error = $_.Exception.Message
          } | ConvertTo-Json -Compress)
        }
      `;

      const result = await this.executePowerShell(script);
      const cleanResult = result.replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\u007F]/g, '').trim();
      const data = JSON.parse(cleanResult);

      if (data.Error) {
        throw new Error(data.Error);
      }

      return data;
    } catch (error) {
      throw new Error(`Failed to get attendee status: ${error instanceof Error ? error.message : String(error)}`);
    }
  }

  /**
   * Get available calendars
   */
  async getCalendars(): Promise<any[]> {
    try {
      const script = `
        try {
          Add-Type -AssemblyName "Microsoft.Office.Interop.Outlook" -ErrorAction Stop
          $outlook = New-Object -ComObject Outlook.Application -ErrorAction Stop
          $namespace = $outlook.GetNamespace("MAPI")
          
          # Build calendars array
          $calendars = @()
          
          # Add default calendar
          $calendars += [PSCustomObject]@{
            Name = "Default"
            Owner = $namespace.CurrentUser.Name
            IsDefault = $true
          }
          
          # Add other calendars
          foreach ($folder in $namespace.Folders) {
            try {
              $calendarFolder = $folder.Folders("Calendar")
              if ($calendarFolder) {
                $calendars += [PSCustomObject]@{
                  Name = "$($folder.Name) - Calendar"
                  Owner = $folder.Name
                  IsDefault = $false
                }
              }
            } catch { }
          }
          
          Write-Output ($calendars | ConvertTo-Json -Compress)
          
        } catch {
          Write-Output ([PSCustomObject]@{
            Error = $_.Exception.Message
          } | ConvertTo-Json -Compress)
        }
      `;

      const result = await this.executePowerShell(script);
      const cleanResult = result.replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\u007F]/g, '').trim();
      const data = JSON.parse(cleanResult);

      if (data.Error) {
        throw new Error(data.Error);
      }

      return Array.isArray(data) ? data : [data];
    } catch (error) {
      throw new Error(`Failed to get calendars: ${error instanceof Error ? error.message : String(error)}`);
    }
  }
}
