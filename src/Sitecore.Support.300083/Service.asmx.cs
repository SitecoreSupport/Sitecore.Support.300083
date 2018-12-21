using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Web.Security;
using System.Web.Services;
using System.Xml;
using Sitecore.Collections;
using Sitecore.Configuration;
using Sitecore.Data;
using Sitecore.Data.Fields;
using Sitecore.Data.Items;
using Sitecore.Data.Managers;
using Sitecore.Data.Masters;
using Sitecore.Diagnostics;
using Sitecore.Globalization;
using Sitecore.Resources;
using Sitecore.Security.Accounts;
using Sitecore.Security.Principal;
using Sitecore.SecurityModel;
using Sitecore.Xml;
using Debug=System.Diagnostics.Debug;
using Version=Sitecore.Data.Version;

namespace Sitecore.Support.Visual {

  ///==========================================================================
  /// <summary></summary>
  ///==========================================================================
  [WebService(Name="Visual Sitecore Service", Namespace="http://sitecore.net/visual/")]
  public class Service : WebService {
    #region Fields

    StringCollection m_data;
    StringCollection m_warnings;
    string m_error;

    #endregion

    ///--------------------------------------------------------------------------
    /// <summary></summary>
    ///--------------------------------------------------------------------------
    public Service() {
      InitializeComponent();
    }

    #region Component Designer generated code

    IContainer components = null;

    ///--------------------------------------------------------------------------
    /// <summary></summary>
    ///--------------------------------------------------------------------------
    void InitializeComponent() {
    }

    ///--------------------------------------------------------------------------
    /// <summary></summary>
    ///--------------------------------------------------------------------------
    protected override void Dispose(bool disposing) {
      if (disposing && components != null) {
        components.Dispose();
      }
      base.Dispose(disposing);
    }

    #endregion

    #region Public methods

    ///--------------------------------------------------------------------------
    /// <summary></summary>
    ///--------------------------------------------------------------------------
    [WebMethod(EnableSession=true)]
    public XmlDocument AddFromMaster(string id, string masterID, string name, string databaseName, Credentials credentials) {
      try {
        Error.AssertString(id, "id", false);
        Error.AssertString(masterID, "masterID", false);
        Error.AssertString(name, "name", false);
        Error.AssertString(databaseName, "databaseName", false);
        Error.AssertObject(credentials, "credentials");

        this.Login(credentials);
        Database database = GetDatabase(databaseName);

        Item item = database.Items[id];
        Debug.Assert(item != null, "Item \"" + id + "\" not found.");
        Debug.Assert(item.Access.CanCreate(), "You do not have permission to create new items at \"" + item.Name + "\".");

        BranchItem branchItem = database.Branches[masterID];
        Debug.Assert(branchItem != null, "Master \"" + masterID + "\" not found.");

        Item newItem = item.Add(name, branchItem);

        AddData(newItem.ID.ToString());
      }
      catch (Exception e) {
        SetError(e);
      }

      return GetResult();
    }

    ///--------------------------------------------------------------------------
    /// <summary></summary>
    ///--------------------------------------------------------------------------
    [WebMethod(EnableSession=true)]
    public XmlDocument AddFromTemplate(string id, string templateID, string name, string databaseName, Credentials credentials) {
      try {
        Error.AssertString(id, "id", false);
        Error.AssertString(templateID, "masterID", false);
        Error.AssertString(name, "name", false);
        Error.AssertString(databaseName, "databaseName", false);
        Error.AssertObject(credentials, "credentials");

        this.Login(credentials);
        Database database = GetDatabase(databaseName);

        Item item = database.Items[id];
        Debug.Assert(item != null, "Item \"" + id + "\" not found.");
        Debug.Assert(item.Access.CanCreate(), "You do not have permission to create new items at \"" + item.Name + "\".");

        TemplateItem templateItem = database.Templates[templateID];
        Debug.Assert(templateItem != null, "Master \"" + templateID + "\" not found.");

        Item newItem = item.Add(name, templateItem);

        AddData(newItem.ID.ToString());
      }
      catch (Exception e) {
        SetError(e);
      }

      return GetResult();
    }

    ///--------------------------------------------------------------------------
    /// <summary></summary>
    ///--------------------------------------------------------------------------
    [WebMethod(EnableSession=true)]
    public XmlDocument AddVersion(string id, string language, string databaseName, Credentials credentials) {
      try {
        Error.AssertString(id, "id", false);
        Error.AssertString(language, "language", false);
        Error.AssertString(databaseName, "databaseName", false);
        Error.AssertObject(credentials, "credentials");

        this.Login(credentials);
        Database database = GetDatabase(databaseName);

        Item item = database.Items[id, Language.Parse(language)];
        Debug.Assert(item != null, "Item \"" + id + "\" not found.");
        Debug.Assert(item.Access.CanWrite(), "You do not have permission to create a new version of \"" + item.Name + "\".");

        Item newItem = item.Versions.AddVersion();

        AddData(newItem.Version.ToString());
      }
      catch (Exception e) {
        SetError(e);
      }

      return GetResult();
    }

    ///--------------------------------------------------------------------------
    /// <summary></summary>
    ///--------------------------------------------------------------------------
    [WebMethod(EnableSession=true)]
    public XmlDocument CopyTo(string id, string newParent, string name, string databaseName, Credentials credentials) {
      try {
        Error.AssertString(id, "id", false);
        Error.AssertString(newParent, "newParent", false);
        Error.AssertString(databaseName, "databaseName", false);
        Error.AssertObject(credentials, "credentials");

        this.Login(credentials);
        Database database = GetDatabase(databaseName);

        Item source = database.Items[id];
        Debug.Assert(source != null, "Item \"" + id + "\" not found.");

        Item parent = database.Items[newParent];
        Debug.Assert(source != null, "Parent item \"" + newParent + "\" not found.");

        Debug.Assert(source.Access.CanCopyTo(parent), "You do not have permission to copy \"" + source.Name + "\" to \"" + parent.Name + "\".");

        source.CopyTo(parent, name);
      }
      catch (Exception e) {
        SetError(e);
      }

      return GetResult();
    }

    ///--------------------------------------------------------------------------
    /// <summary></summary>
    ///--------------------------------------------------------------------------
    [WebMethod(EnableSession=true)]
    public XmlDocument Delete(string id, bool recycle, string databaseName, Credentials credentials) {
      try {
        Error.AssertString(id, "id", false);
        Error.AssertString(databaseName, "databaseName", false);
        Error.AssertObject(credentials, "credentials");

        this.Login(credentials);
        Database database = GetDatabase(databaseName);

        foreach (string itemID in id.Split('|')) {
          if (itemID.Length > 0) {
            try {
              Item item = database.Items[itemID];
              Debug.Assert(item != null, "Item \"" + itemID + "\" not found.");
              Debug.Assert(item.Access.CanDelete(), "You do not have permission to delete \"" + item.Name + "\".");

              if (recycle) {
                item.Recycle();
              }
              else {
                item.Delete();
              }
            }
            catch (Exception e) {
              AddWarning(e);
            }
          }
        }
      }
      catch (Exception e) {
        SetError(e);
      }

      return GetResult();
    }

    ///--------------------------------------------------------------------------
    /// <summary></summary>
    ///--------------------------------------------------------------------------
    [WebMethod(EnableSession=true)]
    public XmlDocument DeleteChildren(string id,string databaseName, Credentials credentials) {
      try {
        Error.AssertString(id, "id", false);
        Error.AssertString(databaseName, "databaseName", false);
        Error.AssertObject(credentials, "credentials");

        this.Login(credentials);
        Database database = GetDatabase(databaseName);

        Item item = database.Items[id];
        Debug.Assert(item != null, "Item \"" + id + "\" not found.");
        Debug.Assert(item.Access.CanDelete(), "You do not have permission to delete \"" + item.Name + "\".");

        item.DeleteChildren();
      }
      catch (Exception e) {
        SetError(e);
      }

      return GetResult();
    }

    ///--------------------------------------------------------------------------
    /// <summary></summary>
    ///--------------------------------------------------------------------------
    [WebMethod(EnableSession=true)]
    public XmlDocument Duplicate(string id, string name, string databaseName, Credentials credentials) {
      try {
        Error.AssertString(id, "id", false);
        Error.AssertString(name, "name", false);
        Error.AssertString(databaseName, "databaseName", false);
        Error.AssertObject(credentials, "credentials");

        this.Login(credentials);
        Database database = GetDatabase(databaseName);

        Item item = database.Items[id];
        Debug.Assert(item != null, "Item \"" + id + "\" not found.");
        Debug.Assert(item.Access.CanDuplicate(), "You do not have permission to duplicate \"" + item.Name + "\".");

        item.Duplicate(name);
      }
      catch (Exception e) {
        SetError(e);
      }

      return GetResult();
    }

    ///--------------------------------------------------------------------------
    /// <summary></summary>
    ///--------------------------------------------------------------------------
    [WebMethod(EnableSession=true)]
    public XmlDocument GetChildren(string id, string databaseName, Credentials credentials) {
      Error.AssertString(id, "id", false);
      Error.AssertString(databaseName, "databaseName", false);
      Error.AssertObject(credentials, "credentials");

      Packet result = new Packet();

      try {
        this.Login(credentials);
        Database database = GetDatabase(databaseName);

        Item item = database.Items[id];

        if (item != null) {
          foreach (Item child in item.Children) {
            result.StartElement("item", "id", child.ID.ToString(), "icon", Themes.MapTheme(child.Appearance.Icon), "haschildren", child.HasChildren ? "1" : "0");
            result.SetValue(child.Name);
            result.EndElement();
          }
        }
      }
      catch {
        // silent
      }

      return result.XmlDocument;
    }

    ///--------------------------------------------------------------------------
    /// <summary></summary>
    ///--------------------------------------------------------------------------
    [WebMethod(EnableSession=true)]
    public XmlDocument GetDatabases(Credentials credentials) {
      Error.AssertObject(credentials, "credentials");

      Packet result = new Packet();

      try {
        Login(credentials);

        foreach(string name in Factory.GetDatabaseNames()) {
          result.AddElement("database", name);
        }
      }
      catch {
        // silent
      }

      return result.XmlDocument;
    }

    ///--------------------------------------------------------------------------
    /// <summary></summary>
    ///--------------------------------------------------------------------------
    [WebMethod(EnableSession=true)]
    public XmlDocument GetItemMasters(string id, string databaseName, Credentials credentials) {
      Error.AssertString(id, "id", false);
      Error.AssertString(databaseName, "databaseName", false);
      Error.AssertObject(credentials, "credentials");

      Packet result = new Packet();

      try {
        this.Login(credentials);
        Database database = GetDatabase(databaseName);

        Item item = database.Items[id];

        if (item != null) {
          List<Item> masters = Masters.GetMasters(item);

          foreach (Item master in masters) {
            result.StartElement("master", "id", master.ID.ToString(), "icon", Themes.MapTheme(master.Appearance.Icon));
            result.SetValue(master.Name);
            result.EndElement();
          }
        }
      }
      catch {
        // silent
      }

      return result.XmlDocument;
    }

    ///--------------------------------------------------------------------------
    /// <summary></summary>
    ///--------------------------------------------------------------------------
    [WebMethod(EnableSession=true)]
    public XmlDocument GetItemFields(string id, string language, string version, bool allFields, string databaseName, Credentials credentials) {
      Error.AssertString(id, "id", false);
      Error.AssertString(language, "language", false);
      Error.AssertString(version, "version", false);
      Error.AssertString(databaseName, "databaseName", false);
      Error.AssertObject(credentials, "credentials");

      Packet result = new Packet();

      try {
        this.Login(credentials);
        Database database = GetDatabase(databaseName);

        Item item = database.Items[id, Language.Parse(language), Version.Parse(version)];

        if (item != null) {
          // fields
          FieldCollection fields  = item.Fields;

          if (allFields) {
            fields.ReadAll();
          }

          fields.Sort();

          foreach(Field field in fields) {
            if (field.Name.Length > 0) {
              result.StartElement("field");

              result.SetAttribute("itemid", item.ID.ToString());
              result.SetAttribute("language", item.Language.ToString());
              result.SetAttribute("version", item.Version.ToString());

              result.SetAttribute("fieldid", field.ID.ToString());
              result.SetAttribute("name", field.Name);
              result.SetAttribute("title", field.Title);
              result.SetAttribute("type", field.Type);
              result.SetAttribute("source", field.Source);
              result.SetAttribute("section", field.Section);
              result.SetAttribute("tooltip", field.ToolTip);

              result.AddElement("value", field.Value);

              result.EndElement();
            }
          }

          // versions
          result.StartElement("versions");

          foreach(Version ver in item.Versions.GetVersionNumbers()) {
            result.AddElement("version", "", "number", ver.ToString());
          }

          result.EndElement();

          // languages
          result.StartElement("languages");

          foreach(Language lang in item.Database.Languages) {
            result.AddElement("language", lang.CultureInfo.DisplayName, "name", lang.Name, "cultureinfo", lang.CultureInfo.Name);
          }

          result.EndElement();

          // path
          result.StartElement("path");
          
          Item current = item;
          while (current != null) {
            result.AddElement("item", "", "id", current.ID.ToString(), "name", current.Name,  "displayname", current.GetUIDisplayName());
            current = current.Parent;
          }

          result.EndElement();
        }
      }
      catch {
        // silent
      }

      return result.XmlDocument;
    }

    ///--------------------------------------------------------------------------
    /// <summary></summary>
    ///--------------------------------------------------------------------------
    [WebMethod(EnableSession=true)]
    public XmlDocument GetLanguages(string databaseName, Credentials credentials) {
      Error.AssertString(databaseName, "databaseName", false);
      Error.AssertObject(credentials, "credentials");

      Packet result = new Packet();

      try {
        this.Login(credentials);
        Database database = GetDatabase(databaseName);

        foreach(Language language in database.Languages) {
          result.AddElement("language", language.CultureInfo.DisplayName, "name", language.Name, "cultureinfo", language.CultureInfo.Name);
        }
      }
      catch {
        // silent
      }

      return result.XmlDocument;
    }

    ///--------------------------------------------------------------------------
    /// <summary></summary>
    ///--------------------------------------------------------------------------
    [WebMethod(EnableSession=true)]
    public XmlDocument GetMasters(string databaseName, Credentials credentials) {
      Error.AssertString(databaseName, "databaseName", false);
      Error.AssertObject(credentials, "credentials");

      Packet result = new Packet();

      try {
        this.Login(credentials);
        Database database = GetDatabase(databaseName);

        Item item = database.Items["/sitecore/templates"];

        foreach (Item master in item.Children) {
          result.StartElement("master", "id", master.ID.ToString(), "icon", Themes.MapTheme(master.Appearance.Icon));
          result.SetValue(master.Name);
          result.EndElement();
        }
      }
      catch {
        // silent
      }

      return result.XmlDocument;
    }

    ///--------------------------------------------------------------------------
    /// <summary></summary>
    ///--------------------------------------------------------------------------
    [WebMethod(EnableSession=true)]
    public XmlDocument GetTemplates(string databaseName, Credentials credentials) {
      Packet result = new Packet();

      try {
        this.Login(credentials);
        Database database = GetDatabase(databaseName);

        TemplateItem[] templates = database.Templates.GetTemplates(LanguageManager.DefaultLanguage);

        foreach (TemplateItem template in templates) {
          result.StartElement("template", "id", template.ID.ToString(), "icon", Themes.MapTheme(template.Icon));
          result.SetValue(template.Name);
          result.EndElement();
        }
      }
      catch {
        // silent
      }

      return result.XmlDocument;
    }

    ///--------------------------------------------------------------------------
    /// <summary></summary>
    ///--------------------------------------------------------------------------
    [WebMethod(EnableSession=true)]
    public XmlDocument GetXML(string id, bool deep, string databaseName, Credentials credentials) 
    {
      try 
      {
        Error.AssertString(id, "id", false);
        Error.AssertString(databaseName, "databaseName", false);
        Error.AssertObject(credentials, "credentials");

        Database database = this.GetDatabase(databaseName);
        this.Login(credentials);

        Item item = database.GetItem(id);
        Debug.Assert(item != null, "Item \"" + id + "\" not found.");
        
        ItemSerializerOptions options = ItemSerializerOptions.GetDefaultOptions();
        options.AllowDefaultValues = false;
        options.AllowStandardValues = false;
        options.ProcessChildren = deep;
        string xml = item.GetOuterXml(options);

        this.AddData(xml);
      }
      catch (Exception e) 
      {
        this.SetError(e);
      }

      return this.GetResult();
    }

    ///--------------------------------------------------------------------------
    /// <summary></summary>
    ///--------------------------------------------------------------------------
    [WebMethod(EnableSession=true)]
    public XmlDocument InsertXML(string id, string xml, bool changeIDs, string databaseName, Credentials credentials) {
      try {
        Error.AssertString(id, "id", false);
        Error.AssertString(databaseName, "databaseName", false);
        Error.AssertObject(credentials, "credentials");

        this.Login(credentials);
        Database database = GetDatabase(databaseName);

        Item item = database.Items[id];
        Debug.Assert(item != null, "Item \"" + id + "\" not found.");

        item.Paste(xml, changeIDs, PasteMode.Overwrite);
      }
      catch (Exception e) {
        SetError(e);
      }

      return GetResult();
    }

    ///--------------------------------------------------------------------------
    /// <summary></summary>
    ///--------------------------------------------------------------------------
    [WebMethod(EnableSession=true)]
    public XmlDocument MoveTo(string id, string newParent, string databaseName, Credentials credentials) {
      try {
        Error.AssertString(id, "id", false);
        Error.AssertString(newParent, "newParent", false);
        Error.AssertString(databaseName, "databaseName", false);
        Error.AssertObject(credentials, "credentials");

        this.Login(credentials);
        Database database = GetDatabase(databaseName);

        Item source = database.Items[id];
        Debug.Assert(source != null, "Item \"" + id + "\" not found.");

        Item parent = database.Items[newParent];
        Debug.Assert(source != null, "Parent item \"" + newParent + "\" not found.");

        Debug.Assert(source.Access.CanMoveTo(parent), "You do not have permission to move \"" + source.Name + "\" to \"" + parent.Name + "\".");

        source.MoveTo(parent);
      }
      catch (Exception e) {
        SetError(e);
      }

      return GetResult();
    }

    ///--------------------------------------------------------------------------
    /// <summary></summary>
    ///--------------------------------------------------------------------------
    [WebMethod(EnableSession=true)]
    public XmlDocument RemoveVersion(string id, string language, string version, string databaseName, Credentials credentials) {
      try {
        Error.AssertString(id, "id", false);
        Error.AssertString(language, "language", false);
        Error.AssertString(version, "version", false);
        Error.AssertString(databaseName, "databaseName", false);
        Error.AssertObject(credentials, "credentials");

        this.Login(credentials);
        Database database = GetDatabase(databaseName);

        Item item = database.Items[id, Language.Parse(language), Version.Parse(version)];
        Debug.Assert(item != null, "Item \"" + id + "\" not found.");
        Debug.Assert(item.Access.CanWrite(), "You do not have permission to delete a new version of \"" + item.Name + "\".");

        item.Versions.RemoveVersion();
      }
      catch (Exception e) {
        SetError(e);
      }

      return GetResult();
    }

    ///--------------------------------------------------------------------------
    /// <summary></summary>
    ///--------------------------------------------------------------------------
    [WebMethod(EnableSession=true)]
    public XmlDocument Rename(string id, string newName, string databaseName, Credentials credentials) {
      try {
        Error.AssertString(id, "id", false);
        Error.AssertString(newName, "newName", false);
        Error.AssertString(databaseName, "databaseName", false);
        Error.AssertObject(credentials, "credentials");

        this.Login(credentials);
        Database database = GetDatabase(databaseName);

        Item item = database.Items[id];
        Debug.Assert(item != null, "Item \"" + id + "\" not found.");
        Debug.Assert(item.Access.CanRename(), "You do not have permission to rename \"" + item.Name + "\".");

          StatisticDisabler statisticDisabler = (item.Versions.Count > 0) ? null : new StatisticDisabler();
          try
          {
              item.Editing.BeginEdit();
              item.Name = newName;
              item.Editing.EndEdit();
             
          }
          finally
          {
              statisticDisabler?.Dispose();
          }
            }
      catch (Exception e) {
        SetError(e);
      }

      return GetResult();
    }

    ///--------------------------------------------------------------------------
    /// <summary></summary>
    ///--------------------------------------------------------------------------
    [WebMethod(EnableSession=true)]
    public XmlDocument Save(string xml, string databaseName, Credentials credentials) {
      try {
        Error.AssertString(xml, "xml", false);
        Error.AssertString(databaseName, "databaseName", false);
        Error.AssertObject(credentials, "credentials");

        this.Login(credentials);
        Database database = GetDatabase(databaseName);

        XmlDocument doc = XmlUtil.LoadXml(xml);

        Hashtable items = new Hashtable();

        foreach(XmlNode node in doc.SelectNodes("/sitecore/field")) {
          string id = XmlUtil.GetAttribute("itemid", node);
          string language = XmlUtil.GetAttribute("language", node);
          string version = XmlUtil.GetAttribute("version", node);

          string fieldid = XmlUtil.GetAttribute("fieldid", node);
          string value = XmlUtil.GetChildValue("value", node);

          string key = id + language + version;

          Item item = items[key] as Item;

          if (item == null) {
            item = database.Items[id, Language.Parse(language), Version.Parse(version)];

            if (item == null) {
              continue;
            }

            items[key] = item;

            item.Editing.BeginEdit();
          }

          Field field = item.Fields[ID.Parse(fieldid)];

          if (field != null) {
            field.Value = value;
          }
        }

        foreach(Item item in items.Values) {
          item.Editing.EndEdit();
        }
      }
      catch (Exception e) {
        SetError(e);
      }

      return GetResult();
    }

    ///--------------------------------------------------------------------------
    /// <summary></summary>
    ///--------------------------------------------------------------------------
    [WebMethod(EnableSession=true)]
    public XmlDocument VerifyCredentials(Credentials credentials) {
      try {
        Error.AssertObject(credentials, "credentials");

        Login(credentials);

        AddData("OK");
      }
      catch (Exception e) {
        SetError(e);
      }

      return GetResult();
    }

    #endregion

    #region Private methods

    ///--------------------------------------------------------------------------
    /// <summary></summary>
    ///--------------------------------------------------------------------------
    void AddData(string data) {
      if (m_data == null) {
        m_data = new StringCollection();
      }

      m_data.Add(data);
    }

    ///--------------------------------------------------------------------------
    /// <summary></summary>
    ///--------------------------------------------------------------------------
    void AddWarning(Exception e) {
      if (m_warnings == null) {
        m_warnings = new StringCollection();
      }

      m_warnings.Add(e.Message);
    }

    ///--------------------------------------------------------------------------
    /// <summary></summary>
    ///--------------------------------------------------------------------------
    void SetError(Exception e) {
      m_error = e.Message;
    }

    ///--------------------------------------------------------------------------
    /// <summary></summary>
    ///--------------------------------------------------------------------------
    Database GetDatabase(string databaseName) {
      Database database = Factory.GetDatabase(databaseName);

      Debug.Assert(database != null, "Database \"" + databaseName + "\" not found.");

      return database;
    }

    ///--------------------------------------------------------------------------
    /// <summary></summary>
    ///--------------------------------------------------------------------------
    XmlDocument GetResult() {
      Packet packet = new Packet();

      if (m_error == null) {
        packet.AddElement("status", "ok");
      }
      else {
        packet.AddElement("status", "failed");
        packet.AddElement("error", m_error);
      }

      if (m_warnings != null) {
        packet.StartElement("warnings");

        foreach (string warning in m_warnings) {
          packet.AddElement("warning", warning);
        }

        packet.EndElement();
      }

      if (m_data != null) {
        packet.StartElement("data");

        foreach (string data in m_data) {
          XmlNode node = packet.AddElement("data", "");
          node.InnerXml = data;
        }

        packet.EndElement();
      }

      return packet.XmlDocument;
    }

    ///--------------------------------------------------------------------------
    /// <summary></summary>
    ///--------------------------------------------------------------------------
    void Login(Credentials credentials) {
      Error.AssertObject(credentials, "credentials");

      if (Sitecore.Context.IsLoggedIn) {
        if(Sitecore.Context.User.Name.Equals(credentials.UserName, StringComparison.OrdinalIgnoreCase)) {
          return;
        }

        Sitecore.Context.Logout();
      }

      bool validated = Membership.ValidateUser(credentials.UserName, credentials.Password);

      Assert.IsTrue(validated, "Unknown username or password.");

      User user = Security.Accounts.User.FromName(credentials.UserName, true);

      UserSwitcher.Enter(user);
    }

    #endregion
  }
}
