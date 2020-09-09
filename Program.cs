using System;
using System.Collections.Generic;
using System.IO;
using System.Json;
using System.Linq;
using System.Reactive.Linq;
using System.Text;
using System.Threading.Tasks;
using CommandLine;
using Microsoft.Exchange.WebServices.Data;
using Microsoft.Extensions.Configuration;
using Task = System.Threading.Tasks.Task;

namespace Exchange_Cache
{
    class Program
    {
        class Options
        {
            [Option('c', "config", HelpText = "Specify the configuration file to use (default: config.json).")]
            public FileInfo Config { get; set; } = new FileInfo("config.json");
        }

        static int Main(string[] args)
        {
            return Parser.Default.ParseArguments<Options>(args)
                .MapResult(
                    options => Run(new ConfigurationBuilder()
                        .AddJsonFile(options.Config.FullName, true)
                        .Build(), options).Result,
                    _ => 1
                );
        }

        static async Task<int> Run(IConfigurationRoot config, Options options)
        {
            var service = new ExchangeService(ExchangeVersion.Exchange2016);
            service.Credentials = new WebCredentials(config["username"], config["password"]);
            service.AutodiscoverUrl(config["email"], redirectionUri => new Uri(redirectionUri).Scheme == "https");

            await LoadFolders(service);
            await GetAllNotJunkMessages(service).ForEachAsync(message => Console.WriteLine(EmailToJson(message).ToString()));

            return 0;
        }

        static Dictionary<string, string> FolderPaths = new Dictionary<string, string>();

        static IObservable<Folder> GetFolders(ExchangeService service)
        {
            return Observable.Create<Folder>(
                async observer =>
                {
                    var foldersView = new FolderView(100)
                    {
                        PropertySet = new PropertySet(BasePropertySet.IdOnly, FolderSchema.ParentFolderId, FolderSchema.DisplayName),
                    };
                    foldersView.Traversal = FolderTraversal.Deep;

                    var emailFolderFilter = new SearchFilter.SearchFilterCollection(LogicalOperator.And)
                    {
                        new SearchFilter.IsEqualTo(FolderSchema.FolderClass, "IPF.Note"),
                    };

                    FindFoldersResults folders;
                    do
                    {
                        folders = await service.FindFolders(WellKnownFolderName.MsgFolderRoot, emailFolderFilter, foldersView);
                        foreach (var item in folders)
                        {
                            observer.OnNext(item as Folder);
                        }
                        foldersView.Offset = folders.NextPageOffset ?? 0;
                    } while (folders.MoreAvailable);

                    observer.OnCompleted();
                }
            );
        }

        static async Task LoadFolders(ExchangeService service)
        {
            var foldersById = new Dictionary<string, Folder>();
            await GetFolders(service).ForEachAsync(folder => foldersById.Add(folder.Id.UniqueId, folder));
            foreach (var folder in foldersById.Values)
            {
                var path = new List<string>();
                var currentFolder = folder;
                while (currentFolder != null)
                {
                    path.Add(currentFolder.DisplayName);
                    if (!foldersById.ContainsKey(currentFolder.ParentFolderId.UniqueId)) break;
                    currentFolder = foldersById[currentFolder.ParentFolderId.UniqueId];
                }
                path.Reverse();
                FolderPaths[folder.Id.UniqueId] = String.Join("/", path);
            }
        }

        static JsonObject EmailToJson(EmailMessage message)
        {
            return new JsonObject(
                new KeyValuePair<string, JsonValue>("id", message.Id.UniqueId),
                new KeyValuePair<string, JsonValue>("folder", FolderPaths[message.ParentFolderId.UniqueId]),
                new KeyValuePair<string, JsonValue>("datetime", message.DateTimeReceived.ToString("O")),
                new KeyValuePair<string, JsonValue>("subject", message.Subject),
                new KeyValuePair<string, JsonValue>("flagged", message.Flag.FlagStatus != ItemFlagStatus.NotFlagged),
                new KeyValuePair<string, JsonValue>("complete", message.Flag.FlagStatus == ItemFlagStatus.Complete),
                message.Flag.FlagStatus == ItemFlagStatus.Complete ? new KeyValuePair<string, JsonValue>("completed", message.Flag.CompleteDate.ToString("O")) : new KeyValuePair<string, JsonValue>("completed", null),
                new KeyValuePair<string, JsonValue>("read", message.IsRead)
            );
        }

        static IObservable<EmailMessage> GetAllNotJunkMessages(ExchangeService service)
        {
            return Observable.Create<EmailMessage>(
                async observer =>
                {
                    // Find Outlook's own search folder "AllItems", which includes all folders in the account.
                    var allItemsView = new FolderView(10);
                    var allItems = await service.FindFolders(WellKnownFolderName.Root,
                        new SearchFilter.IsEqualTo(FolderSchema.DisplayName, "AllItems"),
                        allItemsView);

                    if (allItems.Folders.Count != 1)
                    {
                        throw new MissingMemberException("AllItems");
                    }

                    // Find the Junk folder.
                    var junkFolder = await Folder.Bind(service, WellKnownFolderName.JunkEmail);

                    // Find all items that are flagged and not in the Junk folder.
                    var allFilter = new SearchFilter.SearchFilterCollection(LogicalOperator.And)
                    {
                        new SearchFilter.IsEqualTo(ItemSchema.ItemClass, "IPM.Note"),
                        new SearchFilter.IsNotEqualTo(ItemSchema.ParentFolderId, junkFolder.Id.UniqueId),
                    };
                    var allView = new ItemView(1000)
                    {
                        PropertySet = new PropertySet(BasePropertySet.IdOnly, ItemSchema.ParentFolderId, ItemSchema.DateTimeReceived, ItemSchema.Subject, ItemSchema.Flag, EmailMessageSchema.IsRead),
                    };

                    FindItemsResults<Item> all;
                    do
                    {
                        all = await allItems.Folders[0].FindItems(allFilter, allView);
                        foreach (var item in all.Items)
                        {
                            observer.OnNext(item as EmailMessage);
                        }
                        allView.Offset = all.NextPageOffset ?? 0;
                    } while (all.MoreAvailable);

                    observer.OnCompleted();
                }
            );
        }
    }
}
