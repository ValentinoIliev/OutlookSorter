using System.Reflection;
using System.Text.RegularExpressions;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookSorter.Workers;

public class Azure
{
	private Outlook.Application _outlookApp;
	private Outlook.NameSpace _outlookNamespace;
	private Outlook.MAPIFolder _inbox;
	private List<Outlook.MailItem> _items;
	private Dictionary<int, List<MailItemExtended>> _orderByTask;
	
	public Azure(Outlook.Application app) {
		_items = new List<Outlook.MailItem>();
		_orderByTask = new Dictionary<int, List<MailItemExtended>>();
		_outlookApp = app;
		_outlookNamespace = _outlookApp.GetNamespace("MAPI");
		_inbox = _outlookNamespace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
		GetAllEmails();
		OrderMailsToFolders();
	}

	private void OrderMailsToFolders() {
		foreach (Outlook.MailItem mail in _items) {
			if(mail == null) continue;
			PRWorker wrk = new PRWorker(mail.Subject);
			if (_orderByTask.ContainsKey(wrk.TaskNumber)) {
				_orderByTask[wrk.TaskNumber].Add(new MailItemExtended(mail,wrk));
			}
			else {
				_orderByTask.Add(wrk.TaskNumber,new List<MailItemExtended>());
				_orderByTask[wrk.TaskNumber].Add(new MailItemExtended(mail,wrk));
			}
		}
		
		foreach (var key in _orderByTask.Keys) {
			if (key == -666) {
				
			}
			else {
				var folder = GetFolder(key);
				var tmp =  _orderByTask[key].OrderBy(e => e.SentOn);
				foreach (var item in tmp) {
					try {
						item.Move(folder);
					}
					catch {
						
					}
				}
			}
		}
	}

	private Outlook.Folder GetFolder(int key) {
		Outlook.MAPIFolder parentFolder = _outlookNamespace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox); //inbox

		Outlook.MAPIFolder rootFolder = _outlookNamespace.Folders.GetFirst();

		Outlook.Folder? prFolder = GetOutlookFolderByName(rootFolder, "PR");

		if (prFolder == null) {
			rootFolder.Folders.Add("PR", Outlook.OlDefaultFolders.olFolderInbox);
			prFolder = GetOutlookFolderByName(rootFolder, "PR");
		}
		Outlook.Folder? taskFolder = GetOutlookFolderByName(prFolder, key.ToString());
		if (taskFolder == null) {
			prFolder.Folders.Add(key.ToString(), Outlook.OlDefaultFolders.olFolderInbox);
			taskFolder = taskFolder = GetOutlookFolderByName(prFolder, key.ToString());
		}
		return taskFolder;
	}

	private Outlook.Folder GetOutlookFolderByName(Outlook.MAPIFolder rootFolder, string pr) {
		foreach (Outlook.MAPIFolder folder in rootFolder.Folders)
		{
			if (folder.Name == pr)
			{
				return folder as Outlook.Folder; 
			}
		}
		return null;
	}

	private void GetAllEmails() {
		foreach (var item in _inbox.Items) {
			if (item is Outlook.MailItem email)
			{
				if (email.Sender.Address == "azuredevops@microsoft.com") {
					_items.Add(email);
				}
			}
		}
	}

	internal class PRWorker
	{
		private const string PR_PATTERN = @"(PR) - #(?<TaskNumber>\d+) (?<TaskDescription>.+?) - (?<Project>.+?) (?<Smth>\d+) \((?<ManagerName>.+?)\)";
		private const string TASK_PATTERN = @"(Task) (?<TaskNumber>\d+) - (?<TaskDescription>)";
		private const string USER_STORY_PATTERN = @"(User Story) (?<TaskNumber>\d+) - (?<TaskDescription>)";


		public int TaskNumber { get; private set; } = -666;
		public string TaskDescription { get; private set; }
		public string Project { get; private set; }
		public int Smth { get; private set; }
		public string ManagerName { get; private set; }

		public PRWorker(string input) {
			if (Regex.IsMatch(input ,PR_PATTERN)) {
				Match match = Regex.Match(input, PR_PATTERN);
				PropertyInfo[] properties = typeof(PRWorker).GetProperties(BindingFlags.Instance | BindingFlags.Public);
				foreach (PropertyInfo property in properties) {
					if (match.Groups[property.Name].Success) {
						if (property.PropertyType == typeof(int)) {
							property.SetValue(this, int.Parse(match.Groups[property.Name].Value));
						}
						else {
							property.SetValue(this, match.Groups[property.Name].Value);
						}
					}
				}
			}else if (Regex.IsMatch(input,TASK_PATTERN)) {
				Match match = Regex.Match(input, TASK_PATTERN);
				PropertyInfo[] properties = typeof(PRWorker).GetProperties(BindingFlags.Instance | BindingFlags.Public);
				foreach (PropertyInfo property in properties) {
					if (match.Groups[property.Name].Success) {
						if (property.PropertyType == typeof(int)) {
							property.SetValue(this, int.Parse(match.Groups[property.Name].Value));
						}
						else {
							property.SetValue(this, match.Groups[property.Name].Value);
						}
					}
				}
			}else if (Regex.IsMatch(input,USER_STORY_PATTERN)) {
				Match match = Regex.Match(input, USER_STORY_PATTERN);
				PropertyInfo[] properties = typeof(PRWorker).GetProperties(BindingFlags.Instance | BindingFlags.Public);
				foreach (PropertyInfo property in properties) {
					if (match.Groups[property.Name].Success) {
						if (property.PropertyType == typeof(int)) {
							property.SetValue(this, int.Parse(match.Groups[property.Name].Value));
						}
						else {
							property.SetValue(this, match.Groups[property.Name].Value);
						}
					}
				}
			}
		}
		
		
	}
	
}

internal class MailItemExtended: Outlook.MailItem
{
	public Outlook.MailItem Mail;
	public Azure.PRWorker Properties;

	public MailItemExtended(Outlook.MailItem mail, Azure.PRWorker prWorker) {
		this.Mail = mail;
		this.Properties = prWorker;
	}
	

	void Outlook._MailItem.Close(Outlook.OlInspectorClose SaveMode) {
		Mail.Close(SaveMode);
	}

	public object Copy() {
		return Mail.Copy();
	}

	public void Delete() {
		Mail.Delete();
	}

	public void Display(object Modal = null) {
		Mail.Display(Modal);
	}

	public object Move(Outlook.MAPIFolder DestFldr) {
		return Mail.Move(DestFldr);
	}

	public void PrintOut() {
		Mail.PrintOut();
	}

	public void Save() {
		Mail.Save();
	}

	public void SaveAs(string Path, object Type = null) {
		Mail.SaveAs(Path, Type);
	}

	public void ClearConversationIndex() {
		Mail.ClearConversationIndex();
	}

	Outlook.MailItem Outlook._MailItem.Forward() {
		return Mail.Forward();
	}

	Outlook.MailItem Outlook._MailItem.Reply() {
		return Mail.Reply();
	}

	Outlook.MailItem Outlook._MailItem.ReplyAll() {
		return Mail.ReplyAll();
	}

	void Outlook._MailItem.Send() {
		Mail.Send();
	}

	public void ShowCategoriesDialog() {
		Mail.ShowCategoriesDialog();
	}

	public void AddBusinessCard(Outlook.ContactItem contact) {
		Mail.AddBusinessCard(contact);
	}

	public void MarkAsTask(Outlook.OlMarkInterval MarkInterval) {
		Mail.MarkAsTask(MarkInterval);
	}

	public void ClearTaskFlag() {
		Mail.ClearTaskFlag();
	}

	public Outlook.Conversation GetConversation() {
		return Mail.GetConversation();
	}

	public Outlook.Application Application => Mail.Application;

	public Outlook.OlObjectClass Class => Mail.Class;

	public Outlook.NameSpace Session => Mail.Session;

	public object Parent => Mail.Parent;

	public Outlook.Actions Actions => Mail.Actions;

	public Outlook.Attachments Attachments => Mail.Attachments;

	public string BillingInformation {
		get => Mail.BillingInformation;
		set => Mail.BillingInformation = value;
	}

	public string Body {
		get => Mail.Body;
		set => Mail.Body = value;
	}

	public string Categories {
		get => Mail.Categories;
		set => Mail.Categories = value;
	}

	public string Companies {
		get => Mail.Companies;
		set => Mail.Companies = value;
	}

	public string ConversationIndex => Mail.ConversationIndex;

	public string ConversationTopic => Mail.ConversationTopic;

	public DateTime CreationTime => Mail.CreationTime;

	public string EntryID => Mail.EntryID;

	public Outlook.FormDescription FormDescription => Mail.FormDescription;

	public Outlook.Inspector GetInspector => Mail.GetInspector;

	public Outlook.OlImportance Importance {
		get => Mail.Importance;
		set => Mail.Importance = value;
	}

	public DateTime LastModificationTime => Mail.LastModificationTime;

	public object MAPIOBJECT => Mail.MAPIOBJECT;

	public string MessageClass {
		get => Mail.MessageClass;
		set => Mail.MessageClass = value;
	}

	public string Mileage {
		get => Mail.Mileage;
		set => Mail.Mileage = value;
	}

	public bool NoAging {
		get => Mail.NoAging;
		set => Mail.NoAging = value;
	}

	public int OutlookInternalVersion => Mail.OutlookInternalVersion;

	public string OutlookVersion => Mail.OutlookVersion;

	public bool Saved => Mail.Saved;

	public Outlook.OlSensitivity Sensitivity {
		get => Mail.Sensitivity;
		set => Mail.Sensitivity = value;
	}

	public int Size => Mail.Size;

	public string Subject {
		get => Mail.Subject;
		set => Mail.Subject = value;
	}

	public bool UnRead {
		get => Mail.UnRead;
		set => Mail.UnRead = value;
	}

	public Outlook.UserProperties UserProperties => Mail.UserProperties;

	public bool AlternateRecipientAllowed {
		get => Mail.AlternateRecipientAllowed;
		set => Mail.AlternateRecipientAllowed = value;
	}

	public bool AutoForwarded {
		get => Mail.AutoForwarded;
		set => Mail.AutoForwarded = value;
	}

	public string BCC {
		get => Mail.BCC;
		set => Mail.BCC = value;
	}

	public string CC {
		get => Mail.CC;
		set => Mail.CC = value;
	}

	public DateTime DeferredDeliveryTime {
		get => Mail.DeferredDeliveryTime;
		set => Mail.DeferredDeliveryTime = value;
	}

	public bool DeleteAfterSubmit {
		get => Mail.DeleteAfterSubmit;
		set => Mail.DeleteAfterSubmit = value;
	}

	public DateTime ExpiryTime {
		get => Mail.ExpiryTime;
		set => Mail.ExpiryTime = value;
	}

	public DateTime FlagDueBy {
		get => Mail.FlagDueBy;
		set => Mail.FlagDueBy = value;
	}

	public string FlagRequest {
		get => Mail.FlagRequest;
		set => Mail.FlagRequest = value;
	}

	public Outlook.OlFlagStatus FlagStatus {
		get => Mail.FlagStatus;
		set => Mail.FlagStatus = value;
	}

	public string HTMLBody {
		get => Mail.HTMLBody;
		set => Mail.HTMLBody = value;
	}

	public bool OriginatorDeliveryReportRequested {
		get => Mail.OriginatorDeliveryReportRequested;
		set => Mail.OriginatorDeliveryReportRequested = value;
	}

	public bool ReadReceiptRequested {
		get => Mail.ReadReceiptRequested;
		set => Mail.ReadReceiptRequested = value;
	}

	public string ReceivedByEntryID => Mail.ReceivedByEntryID;

	public string ReceivedByName => Mail.ReceivedByName;

	public string ReceivedOnBehalfOfEntryID => Mail.ReceivedOnBehalfOfEntryID;

	public string ReceivedOnBehalfOfName => Mail.ReceivedOnBehalfOfName;

	public DateTime ReceivedTime => Mail.ReceivedTime;

	public bool RecipientReassignmentProhibited {
		get => Mail.RecipientReassignmentProhibited;
		set => Mail.RecipientReassignmentProhibited = value;
	}

	public Outlook.Recipients Recipients => Mail.Recipients;

	public bool ReminderOverrideDefault {
		get => Mail.ReminderOverrideDefault;
		set => Mail.ReminderOverrideDefault = value;
	}

	public bool ReminderPlaySound {
		get => Mail.ReminderPlaySound;
		set => Mail.ReminderPlaySound = value;
	}

	public bool ReminderSet {
		get => Mail.ReminderSet;
		set => Mail.ReminderSet = value;
	}

	public string ReminderSoundFile {
		get => Mail.ReminderSoundFile;
		set => Mail.ReminderSoundFile = value;
	}

	public DateTime ReminderTime {
		get => Mail.ReminderTime;
		set => Mail.ReminderTime = value;
	}

	public Outlook.OlRemoteStatus RemoteStatus {
		get => Mail.RemoteStatus;
		set => Mail.RemoteStatus = value;
	}

	public string ReplyRecipientNames => Mail.ReplyRecipientNames;

	public Outlook.Recipients ReplyRecipients => Mail.ReplyRecipients;

	public Outlook.MAPIFolder SaveSentMessageFolder {
		get => Mail.SaveSentMessageFolder;
		set => Mail.SaveSentMessageFolder = value;
	}

	public string SenderName => Mail.SenderName;

	public bool Sent => Mail.Sent;

	public DateTime SentOn => Mail.SentOn;

	public string SentOnBehalfOfName {
		get => Mail.SentOnBehalfOfName;
		set => Mail.SentOnBehalfOfName = value;
	}

	public bool Submitted => Mail.Submitted;

	public string To {
		get => Mail.To;
		set => Mail.To = value;
	}

	public string VotingOptions {
		get => Mail.VotingOptions;
		set => Mail.VotingOptions = value;
	}

	public string VotingResponse {
		get => Mail.VotingResponse;
		set => Mail.VotingResponse = value;
	}

	public Outlook.Links Links => Mail.Links;

	public Outlook.ItemProperties ItemProperties => Mail.ItemProperties;

	public Outlook.OlBodyFormat BodyFormat {
		get => Mail.BodyFormat;
		set => Mail.BodyFormat = value;
	}

	public Outlook.OlDownloadState DownloadState => Mail.DownloadState;

	public int InternetCodepage {
		get => Mail.InternetCodepage;
		set => Mail.InternetCodepage = value;
	}

	public Outlook.OlRemoteStatus MarkForDownload {
		get => Mail.MarkForDownload;
		set => Mail.MarkForDownload = value;
	}

	public bool IsConflict => Mail.IsConflict;

	public bool IsIPFax {
		get => Mail.IsIPFax;
		set => Mail.IsIPFax = value;
	}

	public Outlook.OlFlagIcon FlagIcon {
		get => Mail.FlagIcon;
		set => Mail.FlagIcon = value;
	}

	public bool HasCoverSheet {
		get => Mail.HasCoverSheet;
		set => Mail.HasCoverSheet = value;
	}

	public bool AutoResolvedWinner => Mail.AutoResolvedWinner;

	public Outlook.Conflicts Conflicts => Mail.Conflicts;

	public string SenderEmailAddress => Mail.SenderEmailAddress;

	public string SenderEmailType => Mail.SenderEmailType;

	public bool EnableSharedAttachments {
		get => Mail.EnableSharedAttachments;
		set => Mail.EnableSharedAttachments = value;
	}

	public Outlook.OlPermission Permission {
		get => Mail.Permission;
		set => Mail.Permission = value;
	}

	public Outlook.OlPermissionService PermissionService {
		get => Mail.PermissionService;
		set => Mail.PermissionService = value;
	}

	public Outlook.PropertyAccessor PropertyAccessor => Mail.PropertyAccessor;

	public Outlook.Account SendUsingAccount {
		get => Mail.SendUsingAccount;
		set => Mail.SendUsingAccount = value;
	}

	public string TaskSubject {
		get => Mail.TaskSubject;
		set => Mail.TaskSubject = value;
	}

	public DateTime TaskDueDate {
		get => Mail.TaskDueDate;
		set => Mail.TaskDueDate = value;
	}

	public DateTime TaskStartDate {
		get => Mail.TaskStartDate;
		set => Mail.TaskStartDate = value;
	}

	public DateTime TaskCompletedDate {
		get => Mail.TaskCompletedDate;
		set => Mail.TaskCompletedDate = value;
	}

	public DateTime ToDoTaskOrdinal {
		get => Mail.ToDoTaskOrdinal;
		set => Mail.ToDoTaskOrdinal = value;
	}

	public bool IsMarkedAsTask => Mail.IsMarkedAsTask;

	public string ConversationID => Mail.ConversationID;

	public Outlook.AddressEntry Sender {
		get => Mail.Sender;
		set => Mail.Sender = value;
	}

	public string PermissionTemplateGuid {
		get => Mail.PermissionTemplateGuid;
		set => Mail.PermissionTemplateGuid = value;
	}

	public object RTFBody {
		get => Mail.RTFBody;
		set => Mail.RTFBody = value;
	}

	public string RetentionPolicyName => Mail.RetentionPolicyName;

	public DateTime RetentionExpirationDate => Mail.RetentionExpirationDate;

	public event Outlook.ItemEvents_10_OpenEventHandler? Open {
		add => Mail.Open += value;
		remove => Mail.Open -= value;
	}

	public event Outlook.ItemEvents_10_CustomActionEventHandler? CustomAction {
		add => Mail.CustomAction += value;
		remove => Mail.CustomAction -= value;
	}

	public event Outlook.ItemEvents_10_CustomPropertyChangeEventHandler? CustomPropertyChange {
		add => Mail.CustomPropertyChange += value;
		remove => Mail.CustomPropertyChange -= value;
	}

	public event Outlook.ItemEvents_10_ForwardEventHandler? Forward {
		add => ((Microsoft.Office.Interop.Outlook.ItemEvents_10_Event)Mail).Forward += value;
		remove => ((Microsoft.Office.Interop.Outlook.ItemEvents_10_Event)Mail).Forward -= value;
	}

	public event Outlook.ItemEvents_10_CloseEventHandler? Close {
		add => ((Microsoft.Office.Interop.Outlook.ItemEvents_10_Event)Mail).Close += value;
		remove => ((Microsoft.Office.Interop.Outlook.ItemEvents_10_Event)Mail).Close -= value;
	}

	public event Outlook.ItemEvents_10_PropertyChangeEventHandler? PropertyChange {
		add => Mail.PropertyChange += value;
		remove => Mail.PropertyChange -= value;
	}

	public event Outlook.ItemEvents_10_ReadEventHandler? Read {
		add => Mail.Read += value;
		remove => Mail.Read -= value;
	}

	public event Outlook.ItemEvents_10_ReplyEventHandler? Reply {
		add => ((Microsoft.Office.Interop.Outlook.ItemEvents_10_Event)Mail).Reply += value;
		remove => ((Microsoft.Office.Interop.Outlook.ItemEvents_10_Event)Mail).Reply -= value;
	}

	public event Outlook.ItemEvents_10_ReplyAllEventHandler? ReplyAll {
		add => ((Microsoft.Office.Interop.Outlook.ItemEvents_10_Event)Mail).ReplyAll += value;
		remove => ((Microsoft.Office.Interop.Outlook.ItemEvents_10_Event)Mail).ReplyAll -= value;
	}

	public event Outlook.ItemEvents_10_SendEventHandler? Send {
		add => ((Microsoft.Office.Interop.Outlook.ItemEvents_10_Event)Mail).Send += value;
		remove => ((Microsoft.Office.Interop.Outlook.ItemEvents_10_Event)Mail).Send -= value;
	}

	public event Outlook.ItemEvents_10_WriteEventHandler? Write {
		add => Mail.Write += value;
		remove => Mail.Write -= value;
	}

	public event Outlook.ItemEvents_10_BeforeCheckNamesEventHandler? BeforeCheckNames {
		add => Mail.BeforeCheckNames += value;
		remove => Mail.BeforeCheckNames -= value;
	}

	public event Outlook.ItemEvents_10_AttachmentAddEventHandler? AttachmentAdd {
		add => Mail.AttachmentAdd += value;
		remove => Mail.AttachmentAdd -= value;
	}

	public event Outlook.ItemEvents_10_AttachmentReadEventHandler? AttachmentRead {
		add => Mail.AttachmentRead += value;
		remove => Mail.AttachmentRead -= value;
	}

	public event Outlook.ItemEvents_10_BeforeAttachmentSaveEventHandler? BeforeAttachmentSave {
		add => Mail.BeforeAttachmentSave += value;
		remove => Mail.BeforeAttachmentSave -= value;
	}

	public event Outlook.ItemEvents_10_BeforeDeleteEventHandler? BeforeDelete {
		add => Mail.BeforeDelete += value;
		remove => Mail.BeforeDelete -= value;
	}

	public event Outlook.ItemEvents_10_AttachmentRemoveEventHandler? AttachmentRemove {
		add => Mail.AttachmentRemove += value;
		remove => Mail.AttachmentRemove -= value;
	}

	public event Outlook.ItemEvents_10_BeforeAttachmentAddEventHandler? BeforeAttachmentAdd {
		add => Mail.BeforeAttachmentAdd += value;
		remove => Mail.BeforeAttachmentAdd -= value;
	}

	public event Outlook.ItemEvents_10_BeforeAttachmentPreviewEventHandler? BeforeAttachmentPreview {
		add => Mail.BeforeAttachmentPreview += value;
		remove => Mail.BeforeAttachmentPreview -= value;
	}

	public event Outlook.ItemEvents_10_BeforeAttachmentReadEventHandler? BeforeAttachmentRead {
		add => Mail.BeforeAttachmentRead += value;
		remove => Mail.BeforeAttachmentRead -= value;
	}

	public event Outlook.ItemEvents_10_BeforeAttachmentWriteToTempFileEventHandler? BeforeAttachmentWriteToTempFile {
		add => Mail.BeforeAttachmentWriteToTempFile += value;
		remove => Mail.BeforeAttachmentWriteToTempFile -= value;
	}

	public event Outlook.ItemEvents_10_UnloadEventHandler? Unload {
		add => Mail.Unload += value;
		remove => Mail.Unload -= value;
	}

	public event Outlook.ItemEvents_10_BeforeAutoSaveEventHandler? BeforeAutoSave {
		add => Mail.BeforeAutoSave += value;
		remove => Mail.BeforeAutoSave -= value;
	}

	public event Outlook.ItemEvents_10_BeforeReadEventHandler? BeforeRead {
		add => Mail.BeforeRead += value;
		remove => Mail.BeforeRead -= value;
	}

	public event Outlook.ItemEvents_10_AfterWriteEventHandler? AfterWrite {
		add => Mail.AfterWrite += value;
		remove => Mail.AfterWrite -= value;
	}

	public event Outlook.ItemEvents_10_ReadCompleteEventHandler? ReadComplete {
		add => Mail.ReadComplete += value;
		remove => Mail.ReadComplete -= value;
	}
}