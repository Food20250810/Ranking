using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Serialization;

namespace TaiwanPopularDevelopers
{
    public enum RegionType
    {
        Taiwan,
        HongKongAndMacau,
        Malaysia,
        Singapore
    }

    public class RegionConfig
    {
        public RegionType Type { get; set; }
        public string Name { get; set; } = "";
        public string ChineseName { get; set; } = "";
        public string DirectoryName { get; set; } = "";
        public string[] SearchQueries { get; set; } = new string[0];
    }

    public class GitHubUser
    {
        public string Login { get; set; } = "";
        public string Name { get; set; } = "";
        public string Location { get; set; } = "";
        public int Followers { get; set; }
        public int PublicRepos { get; set; }
        public string Bio { get; set; } = "";
        public string AvatarUrl { get; set; } = "";
        public string HtmlUrl { get; set; } = "";
        public DateTime CreatedAt { get; set; }
        public double Score { get; set; }
        public string Type { get; set; } = "";  // "User" 或 "Organization"
        public List<Repository> TopRepositories { get; set; } = new List<Repository>();
        public List<Repository> TopOrganizationRepositories { get; set; } = new List<Repository>();
        public List<Repository> TopContributedRepositories { get; set; } = new List<Repository>();
        public List<Repository> AllRepositories { get; set; } = new List<Repository>();
    }

    public class Repository
    {
        public string Name { get; set; } = "";
        public string FullName { get; set; } = "";
        public int StargazersCount { get; set; }
        public int ForksCount { get; set; }
        public string HtmlUrl { get; set; } = "";
        public string Language { get; set; } = "";
        public bool IsFork { get; set; }
        public string OwnerLogin { get; set; } = "";
        public bool IsOrganization { get; set; }
        
        // 貢獻者排名信息
        public int ContributorRank { get; set; } = 0; // 用戶在此專案中的排名 (1-based)
        public int TotalContributors { get; set; } = 0; // 專案總貢獻者數量
        
        // 格式化顯示排名信息
        public string RankDisplay => ContributorRank > 0 && TotalContributors > 0 
            ? $"(排名{ContributorRank}/{TotalContributors})" 
            : "";
    }

    public class ContributorRankInfo
    {
        public bool IsContributor { get; set; }
        public int Rank { get; set; } = 0; // 排名 (1-based)
        public int TotalContributors { get; set; } = 0; // 總貢獻者數量
        public int CommitCount { get; set; } = 0; // 貢獻的commit數量
    }

    public class RegionProject
    {
        public string Name { get; set; } = "";
        public string FullName { get; set; } = "";
        public int StargazersCount { get; set; }
        public int ForksCount { get; set; }
        public string HtmlUrl { get; set; } = "";
        public string Language { get; set; } = "";
        public string OwnerLogin { get; set; } = "";
        public string OwnerType { get; set; } = ""; // User or Organization
        public string Description { get; set; } = "";
        public string Reason { get; set; } = ""; // 為什麼算是區域專案的原因
        public List<string> RegionContributors { get; set; } = new List<string>(); // 區域貢獻者列表

        /// <summary>
        /// 獲取排序後的區域貢獻者列表，區域開發者排在前面
        /// </summary>
        /// <param name="regionUsers">區域開發者用戶列表</param>
        /// <returns>排序後的貢獻者列表</returns>
        public List<string> GetSortedRegionContributors(List<GitHubUser> regionUsers)
        {
            var regionUserLogins = new HashSet<string>(regionUsers.Select(u => u.Login), StringComparer.OrdinalIgnoreCase);
            
            // 先取區域開發者，再取其他貢獻者，各自按字母順序排序
            var regionContributors = RegionContributors.Where(c => regionUserLogins.Contains(c)).OrderBy(c => c).ToList();
            var otherContributors = RegionContributors.Where(c => !regionUserLogins.Contains(c)).OrderBy(c => c).ToList();
            
            var result = new List<string>();
            result.AddRange(regionContributors);
            result.AddRange(otherContributors);
            
            return result;
        }
    }

    public class GitHubApiResponse<T>
    {
        public T Data { get; set; } = default(T)!;
        public bool IsSuccess { get; set; }
        public string ErrorMessage { get; set; } = "";
        public int RemainingRequests { get; set; }
        public DateTime ResetTime { get; set; }
    }

    public class Program
    {
        private static HttpClient httpClient = new HttpClient();
        private static string? githubToken;
        private static readonly int MinFollowers = 100; // 最低追蹤者數量門檻
        
        // 當前選擇的區域配置
        private static RegionConfig currentRegion = GetTaiwanConfig();
        
        // API調用統計
        private static readonly Dictionary<string, int> apiCallCounts = new Dictionary<string, int>();
        private static readonly Dictionary<string, long> apiCallTimes = new Dictionary<string, long>(); // 毫秒
        private static readonly object statsLock = new object();
        
        // 錯誤統計
        private static readonly Dictionary<string, int> errorCounts = new Dictionary<string, int>();
        private static int totalRetries = 0;
        private static int longWaits = 0;
        
        // 貢獻者排名信息緩存
        private static readonly Dictionary<string, Dictionary<string, ContributorRankInfo>> contributorCache = new Dictionary<string, Dictionary<string, ContributorRankInfo>>();
        private static readonly object cacheLock = new object();
        
        // API響應緩存
        private static readonly Dictionary<string, object> apiResponseCache = new Dictionary<string, object>();
        private static readonly Dictionary<string, DateTime> apiCacheTimestamps = new Dictionary<string, DateTime>();
        private static readonly object apiCacheLock = new object();
        private static readonly TimeSpan cacheExpireTime = TimeSpan.FromHours(24*7); // 緩存24小時過期
        
        // 緩存文件路徑
        private static string cacheDirectory => Path.Combine(currentRegion.DirectoryName, "Cache");
        private static string apiCacheFile => Path.Combine(cacheDirectory, "api_cache.json");
        private static string apiTimestampsFile => Path.Combine(cacheDirectory, "api_timestamps.json");

        /// <summary>
        /// 獲取台灣區域配置
        /// </summary>
        static RegionConfig GetTaiwanConfig()
        {
            return new RegionConfig
            {
                Type = RegionType.Taiwan,
                Name = "Taiwan",
                ChineseName = "台灣",
                DirectoryName = "Taiwan",
                SearchQueries = new string[]
                {
                    $"followers:>{MinFollowers}+location:Taiwan",
                    $"followers:>{MinFollowers}+location:Taipei",
                    $"followers:>{MinFollowers}+location:Kaohsiung",
                    $"followers:>{MinFollowers}+location:\"New Taipei\"",
                    $"followers:>{MinFollowers}+location:Taoyuan",
                    $"followers:>{MinFollowers}+location:Taichung",
                    $"followers:>{MinFollowers}+location:Tainan",
                    $"followers:>{MinFollowers}+location:Hsinchu",
                    $"followers:>{MinFollowers}+location:Keelung",
                    $"followers:>{MinFollowers}+location:Chiayi",
                    $"followers:>{MinFollowers}+location:Changhua",
                    $"followers:>{MinFollowers}+location:Yunlin",
                    $"followers:>{MinFollowers}+location:Nantou",
                    $"followers:>{MinFollowers}+location:Pingtung",
                    $"followers:>{MinFollowers}+location:Yilan",
                    $"followers:>{MinFollowers}+location:Hualien",
                    $"followers:>{MinFollowers}+location:Taitung",
                    $"followers:>{MinFollowers}+location:Penghu",
                    $"followers:>{MinFollowers}+location:Kinmen",
                    $"followers:>{MinFollowers}+location:Matsu"
                }
            };
        }

        /// <summary>
        /// 獲取香港澳門區域配置
        /// </summary>
        static RegionConfig GetHongKongAndMacauConfig()
        {
            return new RegionConfig
            {
                Type = RegionType.HongKongAndMacau,
                Name = "Hong Kong and Macau",
                ChineseName = "香港澳門",
                DirectoryName = "HongKongAndMacau",
                SearchQueries = new string[]
                {
                    // 香港相關地區
                    $"followers:>{MinFollowers}+location:\"Hong Kong\"",
                    $"followers:>{MinFollowers}+location:HK",
                    $"followers:>{MinFollowers}+location:Hongkong",
                    $"followers:>{MinFollowers}+location:\"Hong Kong SAR\"",
                    $"followers:>{MinFollowers}+location:\"香港\"",
                    // 澳門相關地區
                    $"followers:>{MinFollowers}+location:Macau",
                    $"followers:>{MinFollowers}+location:Macao",
                    $"followers:>{MinFollowers}+location:\"Macau SAR\"",
                    $"followers:>{MinFollowers}+location:\"澳門\""
                }
            };
        }

        /// <summary>
        /// 獲取馬來西亞區域配置
        /// </summary>
        static RegionConfig GetMalaysiaConfig()
        {
            return new RegionConfig
            {
                Type = RegionType.Malaysia,
                Name = "Malaysia",
                ChineseName = "馬來西亞",
                DirectoryName = "Malaysia",
                SearchQueries = new string[]
                {
                    // 馬來西亞相關地區
                    $"followers:>{MinFollowers}+location:Malaysia",
                    $"followers:>{MinFollowers}+location:\"Kuala Lumpur\"",
                    $"followers:>{MinFollowers}+location:\"Kuala+Lumpur\"",
                    $"followers:>{MinFollowers}+location:KL",
                    $"followers:>{MinFollowers}+location:Selangor",
                    $"followers:>{MinFollowers}+location:Johor",
                    $"followers:>{MinFollowers}+location:Penang",
                    $"followers:>{MinFollowers}+location:Perak",
                    $"followers:>{MinFollowers}+location:Sabah",
                    $"followers:>{MinFollowers}+location:Sarawak",
                    $"followers:>{MinFollowers}+location:Kedah",
                    $"followers:>{MinFollowers}+location:Kelantan",
                    $"followers:>{MinFollowers}+location:Terengganu",
                    $"followers:>{MinFollowers}+location:Pahang",
                    $"followers:>{MinFollowers}+location:Negeri+Sembilan",
                    $"followers:>{MinFollowers}+location:Melaka",
                    $"followers:>{MinFollowers}+location:Malacca",
                    $"followers:>{MinFollowers}+location:Perlis",
                    $"followers:>{MinFollowers}+location:Putrajaya",
                    $"followers:>{MinFollowers}+location:Labuan"
                }
            };
        }

        /// <summary>
        /// 獲取新加坡區域配置
        /// </summary>
        static RegionConfig GetSingaporeConfig()
        {
            return new RegionConfig
            {
                Type = RegionType.Singapore,
                Name = "Singapore",
                ChineseName = "新加坡",
                DirectoryName = "Singapore",
                SearchQueries = new string[]
                {
                    // 新加坡相關地區
                    $"followers:>{MinFollowers}+location:Singapore",
                    $"followers:>{MinFollowers}+location:SG",
                    $"followers:>{MinFollowers}+location:\"新加坡\"",
                    $"followers:>{MinFollowers}+location:\"Singapore, SG\"",
                    $"followers:>{MinFollowers}+location:\"Singapore, Singapore\""
                }
            };
        }

        private static readonly string[] SearchQueries = {
            $"followers:>{MinFollowers}+location:Taiwan",
           $"followers:>{MinFollowers}+location:Taipei",
           $"followers:>{MinFollowers}+location:Kaohsiung",
           $"followers:>{MinFollowers}+location:\"New Taipei\"",
           $"followers:>{MinFollowers}+location:Taoyuan",
           $"followers:>{MinFollowers}+location:Taichung",
           $"followers:>{MinFollowers}+location:Tainan",
           $"followers:>{MinFollowers}+location:Hsinchu",
           $"followers:>{MinFollowers}+location:Keelung",
           $"followers:>{MinFollowers}+location:Chiayi",
           $"followers:>{MinFollowers}+location:Changhua",
           $"followers:>{MinFollowers}+location:Yunlin",
           $"followers:>{MinFollowers}+location:Nantou",
           $"followers:>{MinFollowers}+location:Pingtung",
           $"followers:>{MinFollowers}+location:Yilan",
           $"followers:>{MinFollowers}+location:Hualien",
           $"followers:>{MinFollowers}+location:Taitung",
           $"followers:>{MinFollowers}+location:Penghu",
           $"followers:>{MinFollowers}+location:Kinmen",
           $"followers:>{MinFollowers}+location:Matsu",
        //    // 香港相關地區
        //    $"followers:>{MinFollowers}+location:\"Hong Kong\"",
        //    $"followers:>{MinFollowers}+location:HK",
        //    $"followers:>{MinFollowers}+location:Hongkong",
        //    $"followers:>{MinFollowers}+location:\"Hong Kong SAR\"",
        //    $"followers:>{MinFollowers}+location:\"香港\"",
        //    // 澳門相關地區
        //    $"followers:>{MinFollowers}+location:Macau",
        //    $"followers:>{MinFollowers}+location:Macao",
        //    $"followers:>{MinFollowers}+location:\"Macau SAR\"",
        //    $"followers:>{MinFollowers}+location:\"澳門\"",
        //    $"followers:>{MinFollowers}+location:\"澳門\""
        };

        /// <summary>
        /// 從緩存獲取API響應
        /// </summary>
        /// <typeparam name="T">響應類型</typeparam>
        /// <param name="cacheKey">緩存鍵</param>
        /// <returns>緩存的響應，如果不存在或已過期則返回null</returns>
        static GitHubApiResponse<T>? GetFromCache<T>(string cacheKey)
        {
            lock (apiCacheLock)
            {
                if (!apiResponseCache.ContainsKey(cacheKey) || !apiCacheTimestamps.ContainsKey(cacheKey))
                    return null;
                
                // 檢查是否過期
                if (DateTime.Now - apiCacheTimestamps[cacheKey] > cacheExpireTime)
                {
                    apiResponseCache.Remove(cacheKey);
                    apiCacheTimestamps.Remove(cacheKey);
                    return null;
                }
                
                var cachedResponse = apiResponseCache[cacheKey] as GitHubApiResponse<T>;
                if (cachedResponse != null)
                {
                    Console.WriteLine($"[緩存命中] {cacheKey}");
                }
                return cachedResponse;
            }
        }
        
        /// <summary>
        /// 將API響應存儲到緩存
        /// </summary>
        /// <typeparam name="T">響應類型</typeparam>
        /// <param name="cacheKey">緩存鍵</param>
        /// <param name="response">API響應</param>
        static void SaveToCache<T>(string cacheKey, GitHubApiResponse<T> response)
        {
            lock (apiCacheLock)
            {
                apiResponseCache[cacheKey] = response;
                apiCacheTimestamps[cacheKey] = DateTime.Now;
                Console.WriteLine($"[緩存保存] {cacheKey}");
            }
        }
        
        /// <summary>
        /// 生成API緩存鍵
        /// </summary>
        /// <param name="url">API URL</param>
        /// <returns>緩存鍵</returns>
        static string GenerateCacheKey(string url)
        {
            // 移除查詢參數中的時間戳等動態參數，保留核心的API路徑和參數
            try
            {
                var uri = new Uri(url);
                var baseUrl = $"{uri.Scheme}://{uri.Host}{uri.AbsolutePath}";
                
                // 保留查詢參數但排序以確保一致性
                if (!string.IsNullOrEmpty(uri.Query))
                {
                    var queryString = uri.Query.TrimStart('?');
                    var queryParams = queryString.Split('&')
                        .Where(param => !string.IsNullOrEmpty(param))
                        .OrderBy(param => param)
                        .ToList();
                    
                    if (queryParams.Any())
                    {
                        baseUrl += "?" + string.Join("&", queryParams);
                    }
                }
                
                return baseUrl;
            }
            catch
            {
                // 如果URL解析失敗，直接返回原URL
                return url;
            }
        }

        /// <summary>
        /// 檢查URL是否應該被緩存
        /// </summary>
        /// <param name="url">API URL</param>
        /// <returns>是否應該緩存</returns>
        static bool ShouldCacheUrl(string url)
        {
            // 不緩存搜索API的結果，因為它們經常變化
            if (url.Contains("/search/"))
                return false;
                
            // 不緩存分頁超過第1頁的請求，因為它們可能經常變化
            if (url.Contains("page=") && !url.Contains("page=1"))
                return false;
                
            // 不緩存stats API，因為GitHub經常返回202狀態
            if (url.Contains("/stats/"))
                return false;
                
            return true;
        }

        /// <summary>
        /// 選擇要處理的區域
        /// </summary>
        static void SelectRegion()
        {
            Console.WriteLine("請選擇要處理的區域:");
            Console.WriteLine("1. 台灣 (Taiwan)");
            Console.WriteLine("2. 香港澳門 (Hong Kong and Macau)");
            Console.WriteLine("3. 馬來西亞 (Malaysia)");
            Console.WriteLine("4. 新加坡 (Singapore)");
            Console.Write("請輸入選擇 (1-4): ");
            
            var choice = Console.ReadLine()?.Trim();
            
            switch (choice)
            {
                case "1":
                    currentRegion = GetTaiwanConfig();
                    break;
                case "2":
                    currentRegion = GetHongKongAndMacauConfig();
                    break;
                case "3":
                    currentRegion = GetMalaysiaConfig();
                    break;
                case "4":
                    currentRegion = GetSingaporeConfig();
                    break;
                default:
                    Console.WriteLine("無效選擇，默認使用台灣區域");
                    currentRegion = GetTaiwanConfig();
                    break;
            }
            
            // 確保目錄存在
            if (!Directory.Exists(currentRegion.DirectoryName))
            {
                Directory.CreateDirectory(currentRegion.DirectoryName);
                Console.WriteLine($"已創建目錄: {currentRegion.DirectoryName}");
            }
            
            // 確保緩存目錄存在
            if (!Directory.Exists(cacheDirectory))
            {
                Directory.CreateDirectory(cacheDirectory);
                Console.WriteLine($"已創建緩存目錄: {cacheDirectory}");
            }
        }

        /// <summary>
        /// 直接生成文件模式：僅使用現有用戶數據生成HTML和Markdown文件
        /// </summary>
        static async Task GenerateFilesOnly()
        {
            try
            {
                Console.WriteLine("正在載入現有用戶資料...");
                var existingUsers = await LoadExistingUsers();

                if (!existingUsers.Any())
                {
                    Console.WriteLine("錯誤：找不到現有的用戶資料文件 (Users.json)");
                    Console.WriteLine("請先運行完整的數據收集程序，或確保 Users.json 文件存在");
                    return;
                }

                Console.WriteLine($"成功載入 {existingUsers.Count} 個用戶資料");

                // 按分數排序，並過濾掉組織用戶
                var rankedUsers = existingUsers
                    .Where(u => u.Type != "Organization")
                    .OrderByDescending(u => u.Score)
                    .ToList();

                Console.WriteLine($"排名包含 {rankedUsers.Count} 個個人用戶");

                // 生成專案排名
                Console.WriteLine($"正在分析{currentRegion.ChineseName}專案...");
                var regionProjects = GenerateRegionProjectsRanking(rankedUsers);
                Console.WriteLine($"找到 {regionProjects.Count} 個{currentRegion.ChineseName}相關專案");

                // 生成用戶排名 Markdown
                Console.WriteLine("正在生成用戶排名 Markdown 文件...");
                var markdown = GenerateMarkdown(rankedUsers);
                var readmePath = Path.Combine(currentRegion.DirectoryName, "README.md");
                await File.WriteAllTextAsync(readmePath, markdown, Encoding.UTF8);
                Console.WriteLine($"✓ {readmePath} 已生成");

                // 生成用戶排名 HTML
                Console.WriteLine("正在生成用戶排名 HTML 文件...");
                var html = GenerateHtml(rankedUsers);
                var indexPath = Path.Combine(currentRegion.DirectoryName, "index.html");
                await File.WriteAllTextAsync(indexPath, html, Encoding.UTF8);
                Console.WriteLine($"✓ {indexPath} 已生成");

                // 生成專案排名 Markdown
                Console.WriteLine($"正在生成{currentRegion.ChineseName}專案排名 Markdown 文件...");
                var projectsMarkdown = await GenerateRegionProjectsMarkdown(regionProjects, rankedUsers);
                var projectsMarkdownPath = Path.Combine(currentRegion.DirectoryName, $"{currentRegion.Name}-Projects.md");
                await File.WriteAllTextAsync(projectsMarkdownPath, projectsMarkdown, Encoding.UTF8);
                Console.WriteLine($"✓ {projectsMarkdownPath} 已生成");

                // 生成專案排名 HTML
                Console.WriteLine($"正在生成{currentRegion.ChineseName}專案排名 HTML 文件...");
                var projectsHtml = await GenerateRegionProjectsHtml(regionProjects, rankedUsers);
                var projectsHtmlPath = Path.Combine(currentRegion.DirectoryName, $"{currentRegion.Name.ToLower()}-projects.html");
                await File.WriteAllTextAsync(projectsHtmlPath, projectsHtml, Encoding.UTF8);
                Console.WriteLine($"✓ {projectsHtmlPath} 已生成");

                Console.WriteLine("\n=== 文件生成完成 ===");
                Console.WriteLine($"更新時間: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
                Console.WriteLine($"用戶總數: {rankedUsers.Count}");
                Console.WriteLine($"{currentRegion.ChineseName}專案總數: {regionProjects.Count}");

                // 顯示前10名用戶
                Console.WriteLine("\n前10名用戶:");
                for (int i = 0; i < Math.Min(10, rankedUsers.Count); i++)
                {
                    var user = rankedUsers[i];
                    Console.WriteLine($"{i + 1,2}. {user.Login} ({user.Name}) - 分數: {user.Score:F0}");
                }

                // 顯示前10個專案
                Console.WriteLine($"\n前10個{currentRegion.ChineseName}專案:");
                for (int i = 0; i < Math.Min(10, regionProjects.Count); i++)
                {
                    var project = regionProjects[i];
                    Console.WriteLine($"{i + 1,2}. {project.Name} - ⭐{project.StargazersCount:N0} (擁有者: {project.OwnerLogin})");
                }

                Console.WriteLine("\n文件生成完成！");
                Console.WriteLine($"• README.md - GitHub用戶排名 README 文件");
                Console.WriteLine($"• index.html - GitHub用戶排名 網頁");
                Console.WriteLine($"• {currentRegion.Name}-Projects.md - {currentRegion.ChineseName}專案排名 README 文件");
                Console.WriteLine($"• {currentRegion.Name.ToLower()}-projects.html - {currentRegion.ChineseName}專案排名 網頁");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"文件生成時發生錯誤: {ex.Message}");
                Console.WriteLine($"詳細錯誤: {ex}");
                throw;
            }
        }

        /// <summary>
        /// 生成區域專案排名列表
        /// </summary>
        /// <param name="users">用戶列表</param>
        /// <returns>區域專案列表</returns>
        static List<RegionProject> GenerateRegionProjectsRanking(List<GitHubUser> users)
        {
            var regionProjects = new Dictionary<string, RegionProject>();
            
            foreach (var user in users)
            {
                // 檢查個人專案（專案擁有者是區域用戶）
                foreach (var repo in user.TopRepositories)
                {
                    if (!regionProjects.ContainsKey(repo.FullName))
                    {
                        regionProjects[repo.FullName] = new RegionProject
                        {
                            Name = repo.Name,
                            FullName = repo.FullName,
                            StargazersCount = repo.StargazersCount,
                            ForksCount = repo.ForksCount,
                            HtmlUrl = repo.HtmlUrl,
                            Language = repo.Language,
                            OwnerLogin = repo.OwnerLogin,
                            OwnerType = "User",
                            Description = "",
                            Reason = $"專案擁有者 {user.Login} 來自{currentRegion.ChineseName}",
                            RegionContributors = new List<string> { user.Login }
                        };
                    }
                    else
                    {
                        // 如果專案已存在，加入區域貢獻者列表
                        if (!regionProjects[repo.FullName].RegionContributors.Contains(user.Login))
                        {
                            regionProjects[repo.FullName].RegionContributors.Add(user.Login);
                        }
                    }
                }
                
                // 檢查組織專案（區域用戶是排名第一的貢獻者）
                foreach (var repo in user.TopOrganizationRepositories)
                {
                    if (repo.ContributorRank == 1) // 只有排名第一的才算
                    {
                        if (!regionProjects.ContainsKey(repo.FullName))
                        {
                            regionProjects[repo.FullName] = new RegionProject
                            {
                                Name = repo.Name,
                                FullName = repo.FullName,
                                StargazersCount = repo.StargazersCount,
                                ForksCount = repo.ForksCount,
                                HtmlUrl = repo.HtmlUrl,
                                Language = repo.Language,
                                OwnerLogin = repo.OwnerLogin,
                                OwnerType = "Organization",
                                Description = "",
                                Reason = $"{currentRegion.ChineseName}開發者 {user.Login} 是專案的第一貢獻者",
                                RegionContributors = new List<string> { user.Login }
                            };
                        }
                        else
                        {
                            if (!regionProjects[repo.FullName].RegionContributors.Contains(user.Login))
                            {
                                regionProjects[repo.FullName].RegionContributors.Add(user.Login);
                            }
                        }
                    }
                }
                
                // 檢查其他個人專案貢獻（區域用戶是排名第一的貢獻者）
                foreach (var repo in user.TopContributedRepositories)
                {
                    if (repo.ContributorRank == 1) // 只有排名第一的才算
                    {
                        if (!regionProjects.ContainsKey(repo.FullName))
                        {
                            regionProjects[repo.FullName] = new RegionProject
                            {
                                Name = repo.Name,
                                FullName = repo.FullName,
                                StargazersCount = repo.StargazersCount,
                                ForksCount = repo.ForksCount,
                                HtmlUrl = repo.HtmlUrl,
                                Language = repo.Language,
                                OwnerLogin = repo.OwnerLogin,
                                OwnerType = "User",
                                Description = "",
                                Reason = $"{currentRegion.ChineseName}開發者 {user.Login} 是專案的第一貢獻者",
                                RegionContributors = new List<string> { user.Login }
                            };
                        }
                        else
                        {
                            if (!regionProjects[repo.FullName].RegionContributors.Contains(user.Login))
                            {
                                regionProjects[repo.FullName].RegionContributors.Add(user.Login);
                            }
                        }
                    }
                }
            }
            
            // 按星星數排序並返回
            return regionProjects.Values
                .OrderByDescending(p => p.StargazersCount)
                .ToList();
        }

      /// <summary>
        /// 生成台灣專案排名的Markdown文檔
        /// </summary>
        /// <param name="projects">台灣專案列表</param>
        /// <param name="taiwanUsers">台灣開發者用戶列表</param>
        /// <returns>Markdown字符串</returns>
        /// <summary>
        /// 生成區域專案排名的Markdown文檔
        /// </summary>
        /// <param name="projects">區域專案列表</param>
        /// <param name="regionUsers">區域開發者用戶列表</param>
        /// <returns>Markdown字符串</returns>
        static async Task<string> GenerateRegionProjectsMarkdown(List<RegionProject> projects, List<GitHubUser> regionUsers)
        {
            var sb = new StringBuilder();
            
            sb.AppendLine($"# {currentRegion.ChineseName}GitHub專案排名");
            sb.AppendLine();
            sb.AppendLine("> 本排名收錄以下類型的專案：");
            sb.AppendLine(">");
            sb.AppendLine("> 1. **個人專案**：專案擁有者來自" + currentRegion.ChineseName);
            sb.AppendLine("> 2. **組織專案**：" + currentRegion.ChineseName + "開發者是該專案的第一貢獻者");
            sb.AppendLine("> 3. **開源貢獻**：" + currentRegion.ChineseName + "開發者是其他專案的第一貢獻者");
            sb.AppendLine(">");
            sb.AppendLine("> 按照 ⭐ Star 數量降序排列，顯示前100名");
            sb.AppendLine();
            sb.AppendLine($"**更新時間**: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
            sb.AppendLine($"**專案總數**: {Math.Min(projects.Count, 100)} (顯示前100名)");
            sb.AppendLine();

            // 取前100名專案
            var top100Projects = projects.Take(100).ToList();

            // 生成表格標題
            sb.AppendLine($"| 排名 | {currentRegion.ChineseName}貢獻者 | 專案名稱 | ⭐ Stars | 🍴 Forks | 語言 | 擁有者 | 原因 |");
            sb.AppendLine("|------|------------|----------|----------|----------|------|--------|------|");

            for (int i = 0; i < top100Projects.Count; i++)
            {
                try{

                    var project = top100Projects[i];
                    var rank = i + 1;
                    var projectName = $"[{project.Name}]({project.HtmlUrl})";
                    var stars = project.StargazersCount.ToString("N0");
                    var forks = project.ForksCount.ToString("N0");
                    var language = string.IsNullOrEmpty(project.Language) ? "-" : project.Language;

                    // 擁有者資訊 (頭像 + 姓名 + 真實姓名)
                    var ownerName = await GetUserDisplayName(project.OwnerLogin, regionUsers);
                    var owner = $"[<img src=\"https://github.com/{project.OwnerLogin}.png&s=32\" width=\"32\" height=\"32\" style=\"border-radius: 50%;\" />](https://github.com/{project.OwnerLogin})<br/>**[{project.OwnerLogin}](https://github.com/{project.OwnerLogin})**";
                    if (!string.IsNullOrEmpty(ownerName) && ownerName != project.OwnerLogin)
                    {
                        owner += $"<br/>{ownerName}";
                    }

                    // 區域貢獻者資訊 (頭像10x10px + 姓名 + 真實姓名)
                    var sortedContributors = project.GetSortedRegionContributors(regionUsers);
                    var contributors = "";
                    if (sortedContributors.Any())
                    {
                        var contributorsList = new List<string>();
                        foreach (var contributor in sortedContributors)
                        {
                            var contributorName = GetUserDisplayNameFromList(contributor, regionUsers);
                            var contributorDisplay = $"[<img src=\"https://github.com/{contributor}.png&s=20\" width=\"10\" height=\"10\" style=\"border-radius: 50%;\" />](https://github.com/{contributor}) **[{contributor}](https://github.com/{contributor})**";
                            if (!string.IsNullOrEmpty(contributorName) && contributorName != contributor)
                            {
                                contributorDisplay += $" ({contributorName})";
                            }
                            contributorsList.Add(contributorDisplay);
                        }
                        contributors = string.Join(" ", contributorsList);
                    }
                    else
                    {
                        contributors = "-";
                    }

                    var reason = project.Reason;

                    // 轉義管道符號以避免表格格式錯誤
                    projectName = projectName.Replace("|", "\\|");
                    language = language.Replace("|", "\\|");
                    reason = reason.Replace("|", "\\|");
                    contributors = contributors.Replace("|", "\\|");
                    owner = owner.Replace("|", "\\|");

                    sb.AppendLine($"| {rank} | {contributors} | {projectName} | {stars} | {forks} | {language} | {owner} | {reason} |");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"生成專案 {i + 1} 時發生錯誤: {ex.Message}");
                    sb.AppendLine($"| {i + 1} | - | - | - | - | - | - | 錯誤: {ex.Message} |");
                }
            }
            
            return sb.ToString();
        }

        /// <summary>
        /// 生成區域專案排名的HTML文檔
        /// </summary>
        /// <param name="projects">區域專案列表</param>
        /// <param name="regionUsers">區域開發者用戶列表</param>
        /// <returns>HTML字符串</returns>
        static async Task<string> GenerateRegionProjectsHtml(List<RegionProject> projects, List<GitHubUser> regionUsers)
        {
            var sb = new StringBuilder();
            sb.AppendLine("<!DOCTYPE html>");
            sb.AppendLine("<html lang=\"zh-Hant\">");
            sb.AppendLine("<head>");
            sb.AppendLine("    <meta charset=\"UTF-8\">");
            sb.AppendLine("    <meta name=\"viewport\" content=\"width=device-width, initial-scale=1.0\">");
            sb.AppendLine($"    <title>{currentRegion.ChineseName}GitHub專案排名</title>");
            sb.AppendLine("    <link rel=\"stylesheet\" href=\"https://cdnjs.cloudflare.com/ajax/libs/normalize/8.0.1/normalize.min.css\">");
            sb.AppendLine("    <style>");
            sb.AppendLine("        body { font-family: 'Segoe UI', 'Noto Sans TC', Arial, sans-serif; background: #f7f7f7; color: #222; margin: 0; padding: 20px; }");
            sb.AppendLine("        .container { max-width: 1200px; margin: 0 auto; }");
            sb.AppendLine("        h1 { text-align: center; margin-bottom: 2rem; color: #333; }");
            sb.AppendLine("        .info { text-align: center; margin-bottom: 2rem; color: #666; }");
            sb.AppendLine("        .stats { text-align: center; margin-bottom: 2rem; font-size: 1.1em; }");
            sb.AppendLine("        table { border-collapse: collapse; width: 100%; background: #fff; box-shadow: 0 2px 8px rgba(0,0,0,0.1); border-radius: 8px; overflow: hidden; }");
            sb.AppendLine("        th, td { padding: 12px 8px; border-bottom: 1px solid #ddd; text-align: left; }");
            sb.AppendLine("        th { background: #2c3e50; color: #fff; font-weight: 600; position: sticky; top: 0; }");
            sb.AppendLine("        tr:hover { background: #f8f9fa; }");
            sb.AppendLine("        .rank { text-align: center; font-weight: bold; color: #e74c3c; }");
            sb.AppendLine("        .project-name { font-weight: 600; }");
            sb.AppendLine("        .project-name a { color: #3498db; text-decoration: none; }");
            sb.AppendLine("        .project-name a:hover { text-decoration: underline; }");
            sb.AppendLine("        .stars, .forks { text-align: right; font-weight: 600; }");
            sb.AppendLine("        .stars { color: #f39c12; }");
            sb.AppendLine("        .forks { color: #27ae60; }");
            sb.AppendLine("        .language { background: #ecf0f1; padding: 4px 8px; border-radius: 4px; font-size: 0.9em; }");
            sb.AppendLine("        .contributors a { color: #9b59b6; text-decoration: none; margin-right: 8px; }");
            sb.AppendLine("        .contributors a:hover { text-decoration: underline; }");
            sb.AppendLine("        .reason { font-style: italic; color: #555; }");
            sb.AppendLine("        .nav-links { text-align: center; margin-bottom: 2rem; }");
            sb.AppendLine("        .nav-links a { color: #3498db; text-decoration: none; margin: 0 15px; padding: 8px 16px; border: 1px solid #3498db; border-radius: 4px; }");
            sb.AppendLine("        .nav-links a:hover { background: #3498db; color: white; }");
            sb.AppendLine("        .avatar { border-radius: 50%; width: 20px; height: 20px; vertical-align: middle; margin-right: 6px; }");
            sb.AppendLine("        .avatar-small { border-radius: 50%; width: 20px; height: 20px; vertical-align: middle; margin-right: 4px; }");
            sb.AppendLine("        .owner-info { display: flex; align-items: center; }");
            sb.AppendLine("        .contributor-item { display: inline-flex; align-items: center; margin-right: 12px; white-space: nowrap; }");
            sb.AppendLine("        .contributor-name { color: #666; font-size: 0.9em; }");
            sb.AppendLine("    </style>");
            sb.AppendLine("</head>");
            sb.AppendLine("<body>");
            sb.AppendLine("    <div class=\"container\">");
            sb.AppendLine($"        <h1>{currentRegion.ChineseName}GitHub專案排名</h1>");
            sb.AppendLine("        <div class=\"nav-links\">");
            sb.AppendLine("            <a href=\"index.html\">🏆 開發者排名</a>");
            var projectsFileName = $"{currentRegion.Name.ToLower()}-projects.html";
            sb.AppendLine($"            <a href=\"{projectsFileName}\">📂 專案排名</a>");
            sb.AppendLine("        </div>");
            
            // 取前100名專案
            var top100Projects = projects.Take(100).ToList();
            
            sb.AppendLine($"        <div class=\"stats\">更新時間: {DateTime.Now:yyyy-MM-dd HH:mm:ss} | 專案總數: {Math.Min(projects.Count, 100)} (顯示前100名)</div>");
            sb.AppendLine("        <div class=\"info\">");
            sb.AppendLine($"            收錄標準：個人專案擁有者來自{currentRegion.ChineseName}，或{currentRegion.ChineseName}開發者是該專案的第一貢獻者<br/>");
            sb.AppendLine("            按照 ⭐ Star 數量降序排列");
            sb.AppendLine("        </div>");
            sb.AppendLine("        <table>");
            sb.AppendLine("            <thead>");
            sb.AppendLine("                <tr>");
            sb.AppendLine("                    <th style=\"width: 60px;\">排名</th>");
            sb.AppendLine("                    <th>專案名稱</th>");
            sb.AppendLine("                    <th style=\"width: 80px;\">⭐ Stars</th>");
            sb.AppendLine("                    <th style=\"width: 80px;\">🍴 Forks</th>");
            sb.AppendLine("                    <th style=\"width: 100px;\">語言</th>");
            sb.AppendLine("                    <th style=\"width: 120px;\">擁有者</th>");
            sb.AppendLine($"                    <th>{currentRegion.ChineseName}貢獻者</th>");
            sb.AppendLine("                    <th>收錄原因</th>");
            sb.AppendLine("                </tr>");
            sb.AppendLine("            </thead>");
            sb.AppendLine("            <tbody>");
            
            for (int i = 0; i < top100Projects.Count; i++)
            {
                var project = top100Projects[i];
                var rank = i + 1;
                var languageDisplay = string.IsNullOrEmpty(project.Language) ? "-" : $"<span class=\"language\">{project.Language}</span>";
                
                // 生成擁有者信息，包含頭像和姓名
                var ownerName = await GetUserDisplayName(project.OwnerLogin, regionUsers);
                var ownerDisplay = project.OwnerLogin;
                if (!string.IsNullOrEmpty(ownerName) && ownerName != project.OwnerLogin)
                {
                    ownerDisplay += $"<br/><span class=\"contributor-name\">{ownerName}</span>";
                }
                var ownerHtml = $"<div class=\"owner-info\"><img class=\"avatar\" src=\"https://github.com/{project.OwnerLogin}.png?size=40\" alt=\"{project.OwnerLogin}\" /><a href=\"https://github.com/{project.OwnerLogin}\" target=\"_blank\">{ownerDisplay}</a></div>";
                
                // 生成區域貢獻者信息，包含10x10px頭像和姓名，區域開發者優先排序
                var sortedContributors = project.GetSortedRegionContributors(regionUsers);
                var contributorsHtml = string.Join(" ", sortedContributors.Select(c => {
                    var contributorName = GetUserDisplayNameFromList(c, regionUsers);
                    var displayName = c;
                    if (!string.IsNullOrEmpty(contributorName) && contributorName != c)
                    {
                        displayName += $"<br/><span class=\"contributor-name\">{contributorName}</span>";
                    }
                    return $"<span class=\"contributor-item\"><img class=\"avatar-small\" src=\"https://github.com/{c}.png?size=20\" alt=\"{c}\" /><a href=\"https://github.com/{c}\" target=\"_blank\">{displayName}</a></span>";
                }));
                
                sb.AppendLine("                <tr>");
                sb.AppendLine($"                    <td class=\"rank\">{rank}</td>");
                sb.AppendLine($"                    <td class=\"project-name\"><a href=\"{project.HtmlUrl}\" target=\"_blank\">{project.Name}</a></td>");
                sb.AppendLine($"                    <td class=\"stars\">{project.StargazersCount:N0}</td>");
                sb.AppendLine($"                    <td class=\"forks\">{project.ForksCount:N0}</td>");
                sb.AppendLine($"                    <td>{languageDisplay}</td>");
                sb.AppendLine($"                    <td>{ownerHtml}</td>");
                sb.AppendLine($"                    <td class=\"contributors\">{contributorsHtml}</td>");
                sb.AppendLine($"                    <td class=\"reason\">{project.Reason}</td>");
                sb.AppendLine("                </tr>");
            }
            
            sb.AppendLine("            </tbody>");
            sb.AppendLine("        </table>");
            sb.AppendLine("    </div>");
            sb.AppendLine("</body>");
            sb.AppendLine("</html>");
            
            return sb.ToString();
        }

        /// <summary>
        /// 獲取用戶的顯示名稱，優先返回真實姓名
        /// </summary>
        /// <param name="username">用戶名</param>
        /// <param name="taiwanUsers">台灣用戶列表</param>
        /// <returns>用戶的顯示名稱</returns>
        static async Task<string> GetUserDisplayName(string username, List<GitHubUser> taiwanUsers)
        {
            // 首先從台灣用戶列表中查找
            var taiwanUser = taiwanUsers.FirstOrDefault(u => u.Login.Equals(username, StringComparison.OrdinalIgnoreCase));
            if (taiwanUser != null && !string.IsNullOrEmpty(taiwanUser.Name) && taiwanUser.Name != taiwanUser.Login)
            {
                return taiwanUser.Name;
            }
            
            // 如果不在台灣用戶列表中，嘗試獲取用戶詳細信息
            try
            {
                return "";
                var userDetail = await GetUserDetail(username);
                if (userDetail != null && !string.IsNullOrEmpty(userDetail.Name) && userDetail.Name != userDetail.Login)
                {
                    return userDetail.Name;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"獲取用戶 {username} 詳細信息時發生錯誤: {ex.Message}");
            }
            
            return "";
        }

        /// <summary>
        /// 從用戶列表中獲取用戶的顯示名稱
        /// </summary>
        /// <param name="username">用戶名</param>
        /// <param name="taiwanUsers">台灣用戶列表</param>
        /// <returns>用戶的顯示名稱</returns>
        static string GetUserDisplayNameFromList(string username, List<GitHubUser> taiwanUsers)
        {
            var user = taiwanUsers.FirstOrDefault(u => u.Login.Equals(username, StringComparison.OrdinalIgnoreCase));
            if (user != null && !string.IsNullOrEmpty(user.Name) && user.Name != user.Login)
            {
                return user.Name;
            }
            return "";
        }
        static async Task<List<GitHubUser>> LoadExistingUsers()
        {
            try
            {
                var userJsonPath = Path.Combine(currentRegion.DirectoryName, "Users.json");
                if (File.Exists(userJsonPath))
                {
                    var jsonContent = await File.ReadAllTextAsync(userJsonPath, Encoding.UTF8);
                    var existingData = JsonConvert.DeserializeObject<dynamic>(jsonContent);
                    if (existingData?.Users != null)
                    {
                        var users = JsonConvert.DeserializeObject<List<GitHubUser>>(existingData.Users.ToString());
                        Console.WriteLine($"載入了 {users.Count} 個已完成的用戶資料");
                        
                        // 過濾掉已知的組織用戶
                        var filteredUsers = new List<GitHubUser>();
                        int removedCount = 0;
                        foreach (var user in users)
                        {
                            if (user.Type != "Organization")
                            {
                                filteredUsers.Add(user);
                            }
                            else
                            {
                                removedCount++;
                            }
                        }
                        
                        if (removedCount > 0)
                        {
                            Console.WriteLine($"從現有資料中移除了 {removedCount} 個組織用戶");
                        }
                        
                        return filteredUsers;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"載入現有用戶資料時發生錯誤: {ex.Message}");
            }
            return new List<GitHubUser>();
        }

        static async Task SaveUserData(List<GitHubUser> users)
        {
            try
            {
                var jsonData = new
                {
                    GeneratedAt = DateTime.Now,
                    TotalUsers = users.Count,
                    Users = users.OrderByDescending(u => u.Score).ToList()
                };
                
                var jsonString = JsonConvert.SerializeObject(jsonData, Formatting.Indented);
                var userJsonPath = Path.Combine(currentRegion.DirectoryName, "Users.json");
                await File.WriteAllTextAsync(userJsonPath, jsonString, Encoding.UTF8);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"儲存用戶資料時發生錯誤: {ex.Message}");
            }
        }

        static async Task Main(string[] args)
        {
            Console.WriteLine("GitHub用戶排名系統");
            Console.WriteLine();
            
            // 選擇區域
            SelectRegion();

            Console.WriteLine($"當前選擇區域: {currentRegion.ChineseName}");
            Console.WriteLine("提醒事項:");
            Console.WriteLine("• 程序運行期間可能會遇到GitHub API臨時錯誤，這是正常現象");
            Console.WriteLine("• 如果看到 'InternalServerError' 或空回應，程序會自動重試");
            Console.WriteLine("• 大型專案的stats/contributors API經常不穩定，程序會自動切換到替代方案");
            Console.WriteLine("• 如果程序卡住很久，可以停止並重新運行，緩存會保留已完成的工作");
            Console.WriteLine("• GitHub API狀態可查看: https://status.github.com/");
            Console.WriteLine("• 使用 --generate 或 -g 參數可直接生成文件而不重新檢索數據");
            Console.WriteLine();

            // 檢查命令行參數
            bool generateOnly = args.Any(arg => arg.ToLower() == "--generate" || arg.ToLower() == "-g");
            
            if (generateOnly)
            {
                Console.WriteLine("直接生成模式：使用現有用戶數據生成HTML和Markdown文件");
                await GenerateFilesOnly();
                return;
            }
            
            Console.WriteLine("按任意鍵繼續，或輸入 '--generate' 進入直接生成模式...");
            var skipResponse = Console.ReadLine()?.ToLower();
            
            // 檢查是否有命令行參數要求直接生成文件
            if (skipResponse != null && (skipResponse.ToLower() == "--generate" || skipResponse.ToLower() == "-g"))
            {
                Console.WriteLine("直接生成模式：使用現有用戶數據生成HTML和Markdown文件");
                await GenerateFilesOnly();
                return;
            }
            
            Console.WriteLine("正在讀取GitHub API Token...");
            Console.WriteLine();

            
            // 加載本地緩存
            await LoadCacheFromFile();
            
            // 清理過期緩存
            CleanExpiredCache();

            // 詢問是否需要跳過用戶直到指定用戶名
            Console.WriteLine("是否需要跳過用戶直到指定用戶名？(y/n)");
            skipResponse = Console.ReadLine()?.ToLower();
            
            string? skipUntilUserName = null;
            
            if (skipResponse == "y" || skipResponse == "yes")
            {
                Console.WriteLine("請輸入要跳過到的用戶名 (Login 或 Name):");
                Console.WriteLine("提示: 您可以先載入現有用戶資料來查看可用的用戶名");
                skipUntilUserName = Console.ReadLine()?.Trim();
                if (!string.IsNullOrEmpty(skipUntilUserName))
                {
                    Console.WriteLine($"將跳過所有用戶直到遇到: {skipUntilUserName}");
                }
                else
                {
                    Console.WriteLine("未輸入用戶名，將正常處理所有用戶");
                    skipUntilUserName = null;
                }
            }

            // 讀取GitHub API Token
            try
            {
                githubToken = await File.ReadAllTextAsync(@"C:\Token");
                githubToken = githubToken.Trim();
                Console.WriteLine("GitHub API Token 已載入");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"警告: 無法讀取GitHub API Token: {ex.Message}");
                Console.WriteLine("將使用匿名API調用（限制較多）");
                githubToken = null;
            }

            Console.WriteLine($"正在搜尋{currentRegion.ChineseName}地區的GitHub用戶...");

            // 設置HttpClient
            SetupHttpClient();

            // 如果有 Token，驗證其有效性
            if (!string.IsNullOrEmpty(githubToken))
            {
            }

            if (!string.IsNullOrEmpty(githubToken))
            {
                Console.WriteLine("已設定 GitHub API 授權");
            }
            else
            {
                Console.WriteLine("使用匿名模式（API 限制較多）");
            }

            var allUsers = new List<GitHubUser>();
            var processedUsers = new HashSet<string>();

            // 載入已完成的用戶資料
            Console.WriteLine("正在載入已完成的用戶資料...");
            var existingUsers = await LoadExistingUsers();
            foreach (var existingUser in existingUsers)
            {
                allUsers.Add(existingUser);
                processedUsers.Add(existingUser.Login);
            }

            // 如果用戶選擇了跳過功能，顯示現有用戶列表供參考
            if (!string.IsNullOrEmpty(skipUntilUserName) && existingUsers.Count > 0)
            {
                Console.WriteLine("\n現有用戶列表 (前20個):");
                var usersToShow = existingUsers.Take(20).ToList();
                for (int i = 0; i < usersToShow.Count; i++)
                {
                    var user = usersToShow[i];
                    Console.WriteLine($"{i + 1,2}. Login: {user.Login,-20} Name: {user.Name}");
                }
                if (existingUsers.Count > 20)
                {
                    Console.WriteLine($"... 還有 {existingUsers.Count - 20} 個用戶");
                }
                
                Console.WriteLine($"\n將尋找目標用戶: {skipUntilUserName}");
                Console.WriteLine("按任意鍵繼續...");
            }

            // 搜尋每個地區的用戶 - 使用當前區域的搜索查詢
            foreach (var query in currentRegion.SearchQueries)
            {
                Console.WriteLine($"搜尋地區: {query}");
                try
                {
                    var users = await SearchGitHubUsers(query);
                    
                    foreach (var user in users)
                    {
                        if (!processedUsers.Contains(user.Login))
                        {
                            processedUsers.Add(user.Login);
                            allUsers.Add(user);
                        }
                    }

                    // 避免API限制，每次搜尋後稍作延遲
                    await Task.Delay(1000);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"搜尋地區 {query} 時發生嚴重錯誤: {ex.Message}");
                    Console.WriteLine("程序即將停止");
                    Environment.Exit(1);
                }
            }

            Console.WriteLine($"找到 {allUsers.Count} 個{currentRegion.ChineseName}地區的GitHub用戶 (其中 {existingUsers.Count} 個已完成)");

            // 重新處理所有用戶的項目信息以獲取最新數據和排名信息
            Console.WriteLine("重新處理所有用戶以獲取最新的項目信息和排名數據...");
            
            bool skipMode = !string.IsNullOrEmpty(skipUntilUserName);
            bool foundTargetUser = false;
            
            for (int i = 0; i < allUsers.Count; i++)
            {
                var user = allUsers[i];
                var isNewUser = existingUsers.All(eu => eu.Login != user.Login);
                var userType = isNewUser ? "新用戶" : "更新用戶";
                
                // 檢查跳過邏輯
                if (skipMode && !foundTargetUser)
                {
                    // 檢查是否找到目標用戶 (比較 Login 和 Name)
                    if (
                        (user.Login != null && user.Login.Equals(skipUntilUserName, StringComparison.OrdinalIgnoreCase)) || 
                        (user.Name != null && user.Name.Equals(skipUntilUserName, StringComparison.OrdinalIgnoreCase))
                    )
                    {
                        foundTargetUser = true;
                        Console.WriteLine($"找到目標用戶: {user.Login} ({user.Name})，開始處理...");
                    }
                    else
                    {
                        Console.WriteLine($"跳過用戶 {i + 1}/{allUsers.Count}: {user.Login} ({user.Name})");
                        continue;
                    }
                }
                
                Console.WriteLine($"處理{userType} {i + 1}/{allUsers.Count}: {user.Login}");
                
                try
                {
                    var shouldKeepUser = await CalculateUserScore(user);
                    
                    if (!shouldKeepUser)
                    {
                        // 如果是組織用戶，從列表中移除
                        allUsers.RemoveAt(i);
                        i--; // 調整索引，因為列表大小改變了
                        Console.WriteLine($"已從列表中移除組織用戶: {user.Login}");
                    }
                    else
                    {
                        Console.WriteLine($"已更新用戶 {user.Login} 的資料");
                    }
                    
                    // 每完成一個用戶就儲存
                    await SaveUserData(allUsers.Where(u => u.Type != "Organization").ToList());
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"處理用戶 {user.Login} 時發生錯誤: {ex.Message}");
                    // 即使某個用戶處理失敗，也要儲存其他已完成的用戶
                    await SaveUserData(allUsers.Where(u => u.Type != "Organization").ToList());
                }
                
                // 避免API限制
                await Task.Delay(500);
            }

            // 如果啟用了跳過模式但沒有找到目標用戶，給出提示
            if (skipMode && !foundTargetUser)
            {
                Console.WriteLine($"警告: 沒有找到目標用戶 '{skipUntilUserName}'，所有用戶都被跳過了");
                Console.WriteLine("可能的原因:");
                Console.WriteLine("1. 用戶名拼寫錯誤");
                Console.WriteLine("2. 該用戶不在當前的用戶列表中");
                Console.WriteLine("3. 用戶名大小寫不匹配");
            }

            // 按分數排序，並最終過濾掉任何組織用戶
            var rankedUsers = allUsers
                .Where(u => u.Type != "Organization")
                .OrderByDescending(u => u.Score)
                .ToList();
                
            Console.WriteLine($"最終排名包含 {rankedUsers.Count} 個個人用戶");

            // 最終儲存到JSON檔案
            await SaveUserData(rankedUsers);
            Console.WriteLine("所有用戶資料已最終儲存到 Users.json");

            // 生成專案排名
            Console.WriteLine($"正在分析{currentRegion.ChineseName}專案...");
            var regionProjects = GenerateRegionProjectsRanking(rankedUsers);
            Console.WriteLine($"找到 {regionProjects.Count} 個{currentRegion.ChineseName}相關專案");

            // 生成用戶排名 Markdown
            var markdown = GenerateMarkdown(rankedUsers);
            var readmePath = Path.Combine(currentRegion.DirectoryName, "README.md");
            await File.WriteAllTextAsync(readmePath, markdown, Encoding.UTF8);
            
            // 生成用戶排名 HTML
            var html = GenerateHtml(rankedUsers);
            var indexPath = Path.Combine(currentRegion.DirectoryName, "index.html");
            await File.WriteAllTextAsync(indexPath, html, Encoding.UTF8);
            
            // 生成專案排名 Markdown
            var projectsMarkdown = await GenerateRegionProjectsMarkdown(regionProjects, rankedUsers);
            var projectsMarkdownPath = Path.Combine(currentRegion.DirectoryName, $"{currentRegion.Name}-Projects.md");
            await File.WriteAllTextAsync(projectsMarkdownPath, projectsMarkdown, Encoding.UTF8);
            
            // 生成專案排名 HTML
            var projectsHtml = await GenerateRegionProjectsHtml(regionProjects, rankedUsers);
            var projectsHtmlPath = Path.Combine(currentRegion.DirectoryName, $"{currentRegion.Name.ToLower()}-projects.html");
            await File.WriteAllTextAsync(projectsHtmlPath, projectsHtml, Encoding.UTF8);
            
            Console.WriteLine("排名已生成並儲存到以下文件:");
            Console.WriteLine($"• {readmePath} - GitHub用戶排名 README 文件");
            Console.WriteLine($"• {indexPath} - GitHub用戶排名 網頁");
            Console.WriteLine($"• {projectsMarkdownPath} - {currentRegion.ChineseName}專案排名 README 文件");
            Console.WriteLine($"• {projectsHtmlPath} - {currentRegion.ChineseName}專案排名 網頁");
            
            Console.WriteLine($"\n前10名用戶:");
            for (int i = 0; i < Math.Min(10, rankedUsers.Count); i++)
            {
                var user = rankedUsers[i];
                Console.WriteLine($"{i + 1}. {user.Name} (@{user.Login}) - 分數: {user.Score:F0}");
            }
            
            Console.WriteLine($"\n前10個{currentRegion.ChineseName}專案:");
            for (int i = 0; i < Math.Min(10, regionProjects.Count); i++)
            {
                var project = regionProjects[i];
                Console.WriteLine($"{i + 1}. {project.Name} - ⭐{project.StargazersCount:N0} (擁有者: {project.OwnerLogin})");
            }
            
            // 打印API調用統計摘要
            PrintApiStatsSummary();
            
            // 打印緩存統計
            PrintCacheStats();
            
            // 保存緩存到文件
            await SaveCacheToFile();
            
            Console.WriteLine("\n程序執行完成，緩存已保存到本地文件。");
        }

        static async Task<List<GitHubUser>> SearchGitHubUsers(string query)
        {
            var users = new List<GitHubUser>();
            int page = 1;
            const int maxPages = 1000; // 每個查詢最多100頁
            bool hasUsersWith50PlusFollowers = true;

            while (page <= maxPages && hasUsersWith50PlusFollowers)
            {
                //增加 url follower > 最小追蹤數量
                var url = $"https://api.github.com/search/users?q={query}&sort=followers&order=desc&page={page}&per_page=100";
                
                var response = await MakeGitHubApiCall<dynamic>(url);
                
                if (!response.IsSuccess)
                {
                    Console.WriteLine($"搜尋用戶時發生錯誤: {response.ErrorMessage}");
                    
                    // 如果是服務不可用錯誤，程序已經在 MakeGitHubApiCall 中處理並退出
                    // 如果是其他錯誤，跳出循環但不終止程序
                    break;
                }

                var items = response.Data?.items;
                if (items == null)
                    break;
                    
                if (items.Count == 0)
                    break;

                hasUsersWith50PlusFollowers = false;
                foreach (var item in items)
                {
                    var followers = item.followers ?? 0;
                    var user = new GitHubUser
                    {
                        Login = item.login,
                        Name = item.name ?? item.login,
                        Location = item.location ?? "",
                        Followers = followers,
                        PublicRepos = item.public_repos ?? 0,
                        AvatarUrl = item.avatar_url ?? "",
                        HtmlUrl = item.html_url ?? ""
                    };
                    users.Add(user);
                    //if (followers >= MinFollowers)
                    //{
                        hasUsersWith50PlusFollowers = true;
                    //}
                }

                page++;
                await Task.Delay(100); // 避免API限制
            }

            return users;
        }

        static async Task<bool> CalculateUserScore(GitHubUser user)
        {
            // 獲取用戶詳細資訊
            var userDetail = await GetUserDetail(user.Login);
            if (userDetail != null)
            {
                user.Name = userDetail.Name;
                user.Bio = userDetail.Bio;
                user.CreatedAt = userDetail.CreatedAt;
                user.Followers = userDetail.Followers;  // 更新 Followers
                user.PublicRepos = userDetail.PublicRepos;  // 更新 PublicRepos
                user.Location = userDetail.Location;  // 也更新 Location 以確保準確性
                user.Type = userDetail.Type;  // 更新用戶類型
                
                // 如果是組織，返回 false 表示應該被排除
                if (user.Type == "Organization")
                {
                    Console.WriteLine($"跳過組織用戶: {user.Login}");
                    return false;
                }
            }

            // 獲取用戶的所有個人倉庫
            var personalRepos = await GetAllUserRepositories(user.Login);
            
            // 獲取用戶參與的組織倉庫
            var orgRepos = await GetUserOrganizationRepositories(user.Login);
            
            // 獲取用戶貢獻的其他個人專案（非自己的且非組織的）
            var contributedRepos = await GetUserContributedRepositories(user.Login);
            
            // 合併所有類型的專案到 AllRepositories
            var allRepositories = new List<Repository>();
            allRepositories.AddRange(personalRepos);
            allRepositories.AddRange(orgRepos);
            allRepositories.AddRange(contributedRepos);
            user.AllRepositories = allRepositories;

            // 獲取用戶的頂級個人倉庫（前五名）
            var topRepos = personalRepos.OrderByDescending(r => r.StargazersCount + r.ForksCount).Take(5).ToList();
            user.TopRepositories = topRepos;

            // 獲取頂級組織貢獻專案（前五名）
            var topOrgRepos = orgRepos.OrderByDescending(r => r.StargazersCount + r.ForksCount).Take(5).ToList();
            user.TopOrganizationRepositories = topOrgRepos;

            // 獲取頂級其他個人專案貢獻（前五名）
            var topContributedRepos = contributedRepos.OrderByDescending(r => r.StargazersCount + r.ForksCount).Take(5).ToList();
            user.TopContributedRepositories = topContributedRepos;

            // 計算分數
            double score = 0;
            
            // 個人追蹤數量
            score += user.Followers * 1.0;
            
            // 個人專案star + fork (保持原有計算方式)
            var personalProjectScore = user.TopRepositories.Sum(r => r.StargazersCount * 1.0 + r.ForksCount * 1.0);
            score += personalProjectScore;
            
            // 組織貢獻專案分數 - 使用基於排名的積分計算
            var orgProjectScore = CalculateRankBasedScore(user.TopOrganizationRepositories);
            score += orgProjectScore;
            
            // 其他個人專案貢獻分數 - 使用基於排名的積分計算
            var contributedProjectScore = CalculateRankBasedScore(user.TopContributedRepositories);
            score += contributedProjectScore;

            // 檢查是否滿足最低要求：除了追蹤者之外的其他分數加起來要超過 10
            var otherScores = personalProjectScore + orgProjectScore + contributedProjectScore;
            if (otherScores <= 10)
            {
                Console.WriteLine($"跳過用戶 {user.Login}: 除追蹤者外的其他分數只有 {otherScores:F0}，低於最低要求 10");
                Console.WriteLine($"  個人專案分數: {personalProjectScore:F0}");
                Console.WriteLine($"  組織專案分數: {orgProjectScore:F0}");
                Console.WriteLine($"  其他專案分數: {contributedProjectScore:F0}");
                return false; // 返回 false 表示用戶應該被排除
            }

            Console.WriteLine($"用戶 {user.Login} 總分數: {score:F0} (追蹤者: {user.Followers}, 個人專案: {personalProjectScore:F0}, 組織專案: {orgProjectScore:F0}, 其他專案: {contributedProjectScore:F0})");
            user.Score = score;
            return true; // 返回 true 表示用戶應該被保留
        }

        /// <summary>
        /// 基於貢獻者排名百分比計算項目分數
        /// 新公式: 分別計算 star 積分和 fork 積分
        /// Star積分 = ((總貢獻者數 - 排名 + 1) / 總貢獻者數) * 0.01 * star
        /// Fork積分 = ((總貢獻者數 - 排名 + 1) / 總貢獻者數) * 0.01 * fork
        /// 第1名: (總數/總數) * 0.01 * star/fork = 1.0 * 0.01 * star/fork
        /// 第2名: ((總數-1)/總數) * 0.01 * star/fork
        /// 第n名: ((總數-n+1)/總數) * 0.01 * star/fork
        /// </summary>
        /// <param name="repositories">倉庫列表</param>
        /// <returns>計算後的總分數</returns>
        static double CalculateRankBasedScore(List<Repository> repositories)
        {
            double totalScore = 0;
            
            foreach (var repo in repositories)
            {
                if (repo.ContributorRank > 0 && repo.TotalContributors > 0)
                {
                    // 計算排名百分比: (總貢獻者數 - 排名 + 1) / 總貢獻者數
                    var rankPercentage = (double)(repo.TotalContributors - repo.ContributorRank + 1) / repo.TotalContributors;
                    
                    // 分別計算 star 積分和 fork 積分
                    var starScore = rankPercentage * repo.StargazersCount;
                    var forkScore = rankPercentage * repo.ForksCount;
                    
                    // 總分數
                    var repoScore = starScore + forkScore;
                    totalScore += repoScore;
                    
                    Console.WriteLine($"    {repo.FullName}: 排名{repo.ContributorRank}/{repo.TotalContributors} (百分比: {rankPercentage:P1})");
                    Console.WriteLine($"      Star積分: {rankPercentage:F3} * {repo.StargazersCount} = {starScore:F2}");
                    Console.WriteLine($"      Fork積分: {rankPercentage:F3} * {repo.ForksCount} = {forkScore:F2}");
                    Console.WriteLine($"      總積分: {repoScore:F2}");
                }
                else
                {
                    // 如果沒有排名信息，使用傳統計算方式
                    var fallbackScore = repo.StargazersCount + repo.ForksCount;
                    totalScore += fallbackScore;
                    Console.WriteLine($"    {repo.FullName}: 無排名信息，使用傳統計算 = {fallbackScore:F0}");
                }
            }
            
            return totalScore;
        }

        // 從文件加載緩存
        static async Task LoadCacheFromFile()
        {
            try
            {
                // 確保緩存目錄存在
                if (!Directory.Exists(cacheDirectory))
                {
                    Directory.CreateDirectory(cacheDirectory);
                    Console.WriteLine($"建立緩存目錄: {cacheDirectory}");
                    return;
                }

                // 加載API緩存
                if (File.Exists(apiCacheFile))
                {
                    var cacheJson = await File.ReadAllTextAsync(apiCacheFile, Encoding.UTF8);
                    var loadedCache = JsonConvert.DeserializeObject<Dictionary<string, object>>(cacheJson);
                    
                    if (loadedCache != null)
                    {
                        lock (apiCacheLock)
                        {
                            foreach (var kvp in loadedCache)
                            {
                                apiResponseCache[kvp.Key] = kvp.Value;
                            }
                        }
                        Console.WriteLine($"載入 {loadedCache.Count} 個API緩存項目");
                    }
                }

                // 加載時間戳緩存
                if (File.Exists(apiTimestampsFile))
                {
                    var timestampsJson = await File.ReadAllTextAsync(apiTimestampsFile, Encoding.UTF8);
                    var loadedTimestamps = JsonConvert.DeserializeObject<Dictionary<string, DateTime>>(timestampsJson);
                    
                    if (loadedTimestamps != null)
                    {
                        lock (apiCacheLock)
                        {
                            foreach (var kvp in loadedTimestamps)
                            {
                                apiCacheTimestamps[kvp.Key] = kvp.Value;
                            }
                        }
                        Console.WriteLine($"載入 {loadedTimestamps.Count} 個時間戳緩存項目");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"載入緩存時發生錯誤: {ex.Message}");
            }
        }

        // 保存緩存到文件
        static async Task SaveCacheToFile()
        {
            try
            {
                // 確保緩存目錄存在
                if (!Directory.Exists(cacheDirectory))
                {
                    Directory.CreateDirectory(cacheDirectory);
                }

                // 保存API緩存
                Dictionary<string, object> cacheToSave;
                Dictionary<string, DateTime> timestampsToSave;
                
                lock (apiCacheLock)
                {
                    cacheToSave = new Dictionary<string, object>(apiResponseCache);
                    timestampsToSave = new Dictionary<string, DateTime>(apiCacheTimestamps);
                }

                var cacheJson = JsonConvert.SerializeObject(cacheToSave, Formatting.Indented);
                await File.WriteAllTextAsync(apiCacheFile, cacheJson, Encoding.UTF8);
                Console.WriteLine($"保存 {cacheToSave.Count} 個API緩存項目到文件");

                // 保存時間戳緩存
                var timestampsJson = JsonConvert.SerializeObject(timestampsToSave, Formatting.Indented);
                await File.WriteAllTextAsync(apiTimestampsFile, timestampsJson, Encoding.UTF8);
                Console.WriteLine($"保存 {timestampsToSave.Count} 個時間戳緩存項目到文件");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"保存緩存時發生錯誤: {ex.Message}");
            }
        }

        static async Task<GitHubUser?> GetUserDetail(string username)
        {
            var url = $"https://api.github.com/users/{username}";
            var response = await MakeGitHubApiCall<dynamic>(url);
            
            if (!response.IsSuccess)
                return null;

            return new GitHubUser
            {
                Login = response.Data.login,
                Name = response.Data.name ?? response.Data.login,
                Location = response.Data.location ?? "",
                Followers = response.Data.followers ?? 0,
                PublicRepos = response.Data.public_repos ?? 0,
                Bio = response.Data.bio ?? "",
                AvatarUrl = response.Data.avatar_url ?? "",
                HtmlUrl = response.Data.html_url ?? "",
                Type = response.Data.type ?? "User",
                CreatedAt = response.Data.created_at != null ? DateTime.Parse(response.Data.created_at.ToString()) : DateTime.MinValue
            };
        }

        /// <summary>
        /// 獲取用戶在指定倉庫中的貢獻者排名信息
        /// </summary>
        /// <param name="username">用戶名</param>
        /// <param name="repoFullName">倉庫全名</param>
        /// <returns>貢獻者排名信息</returns>
        static async Task<ContributorRankInfo> GetContributorRankInfo(string username, string repoFullName)
        {
            // 檢查緩存
            lock (cacheLock)
            {
                if (contributorCache.ContainsKey(repoFullName) && contributorCache[repoFullName].ContainsKey(username))
                {
                    Console.WriteLine($"從緩存獲取 {username} 在 {repoFullName} 的排名信息");
                    return contributorCache[repoFullName][username];
                }
            }
            
            var result = new ContributorRankInfo();
            
            try
            {
                // 首先嘗試使用 stats/contributors API 獲取詳細的貢獻統計
                var statsUrl = $"https://api.github.com/repos/{repoFullName}/stats/contributors";
                Console.WriteLine($"正在獲取 {username} 在 {repoFullName} 的詳細統計信息...");
                var statsResponse = await MakeGitHubApiCall<List<dynamic>>(statsUrl);
                
                // 檢查 stats/contributors API 是否成功返回有效資料
                if (statsResponse.IsSuccess && statsResponse.Data != null && statsResponse.Data.Count > 0)
                {
                    Console.WriteLine($"成功獲取 {repoFullName} 的詳細統計數據，包含 {statsResponse.Data.Count} 個貢獻者");
                    var contributorStats = new List<(string Login, int Commits)>();
                    
                    foreach (var contributor in statsResponse.Data)
                    {
                        var login = contributor.author?.login?.ToString();
                        if (string.IsNullOrEmpty(login)) continue;
                        
                        // 計算總提交數
                        int totalCommits = SafeGetInt(contributor.total, 0);
                        if (totalCommits > 0)
                        {
                            contributorStats.Add((login!, totalCommits));
                        }
                    }
                    
                    // 按提交數排序
                    var sortedContributors = contributorStats.OrderByDescending(c => c.Commits).ToList();
                    result.TotalContributors = sortedContributors.Count;
                    
                    // 查找用戶排名
                    for (int i = 0; i < sortedContributors.Count; i++)
                    {
                        if (sortedContributors[i].Login == username)
                        {
                            result.IsContributor = true;
                            result.Rank = i + 1; // 1-based ranking
                            result.CommitCount = sortedContributors[i].Commits;
                            break;
                        }
                    }
                    
                    return result;
                }
                else
                {
                    // Stats API 沒有返回有效數據，記錄詳細信息
                    if (!statsResponse.IsSuccess)
                    {
                        Console.WriteLine($"Stats API 調用失敗 - {repoFullName}:");
                        Console.WriteLine($"  錯誤信息: {statsResponse.ErrorMessage}");
                        Console.WriteLine($"  可能原因: GitHub伺服器問題、專案過大、或API臨時不可用");
                    }
                    else if (statsResponse.Data == null)
                    {
                        Console.WriteLine($"Stats API 返回空數據 - {repoFullName}");
                        Console.WriteLine($"  可能原因: 專案沒有貢獻者統計或數據正在計算中");
                    }
                    else
                    {
                        Console.WriteLine($"Stats API 返回空列表 - {repoFullName}");
                        Console.WriteLine($"  返回數據數量: {statsResponse.Data.Count}");
                    }
                    Console.WriteLine($"  切換到基本 contributors API 作為替代方案...");
                }
                
                // 如果 stats API 失敗，回退到基本的 contributors API
                var contributorsUrl = $"https://api.github.com/repos/{repoFullName}/contributors?per_page=100";
                var page = 1;
                var allContributors = new List<(string Login, int Contributions)>();
                
                while (page <= 10) // 最多查10頁，避免無限循環
                {
                    var pagedUrl = $"{contributorsUrl}&page={page}";
                    var contributorsResponse = await MakeGitHubApiCall<List<dynamic>>(pagedUrl);
                    
                    if (!contributorsResponse.IsSuccess || contributorsResponse.Data == null || contributorsResponse.Data.Count == 0)
                        break;
                    
                    foreach (var contributor in contributorsResponse.Data)
                    {
                        var login = contributor.login?.ToString();
                        var contributions = SafeGetInt(contributor.contributions, 0);
                        
                        if (!string.IsNullOrEmpty(login) && contributions > 0)
                        {
                            allContributors.Add((login!, contributions));
                        }
                    }
                    
                    page++;
                    await Task.Delay(50); // 短暫延遲
                }
                
                if (allContributors.Count > 0)
                {
                    // 按貢獻數排序
                    var sortedContributors = allContributors.OrderByDescending(c => c.Contributions).ToList();
                    result.TotalContributors = sortedContributors.Count;
                    
                    // 查找用戶排名
                    for (int i = 0; i < sortedContributors.Count; i++)
                    {
                        if (sortedContributors[i].Login == username)
                        {
                            result.IsContributor = true;
                            result.Rank = i + 1; // 1-based ranking
                            result.CommitCount = sortedContributors[i].Contributions;
                            break;
                        }
                    }
                }
                
                // 保存到緩存
                lock (cacheLock)
                {
                    if (!contributorCache.ContainsKey(repoFullName))
                    {
                        contributorCache[repoFullName] = new Dictionary<string, ContributorRankInfo>();
                    }
                    contributorCache[repoFullName][username] = result;
                }
                
                return result;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"獲取用戶 {username} 在 {repoFullName} 的排名信息時發生異常: {ex.Message}");
                return result;
            }
        }

        /// <summary>
        /// 檢查用戶是否為倉庫貢獻者（向後兼容的方法）
        /// </summary>
        static async Task<bool> IsUserTopContributor(string username, string repoFullName, int topCount = 10)
        {
            var rankInfo = await GetContributorRankInfo(username, repoFullName);
            return rankInfo.IsContributor; // 只要有貢獻就返回true
        }

        /// <summary>
        /// 安全地從動態對象中獲取布爾值
        /// </summary>
        /// <param name="value">動態對象的值</param>
        /// <param name="defaultValue">默認值</param>
        /// <returns>轉換後的布爾值</returns>
        static bool SafeGetBool(dynamic value, bool defaultValue = false)
        {
            if (value == null) return defaultValue;
            
            if (value is bool boolValue)
                return boolValue;
                
            if (bool.TryParse(value.ToString(), out bool parsedValue))
                return parsedValue;
                
            return defaultValue;
        }

        /// <summary>
        /// 安全地從動態對象中獲取整數值
        /// </summary>
        /// <param name="value">動態對象的值</param>
        /// <param name="defaultValue">默認值</param>
        /// <returns>轉換後的整數值</returns>
        static int SafeGetInt(dynamic value, int defaultValue = 0)
        {
            if (value == null) return defaultValue;
            
            if (value is int intValue)
                return intValue;
            
            if (value is long longValue)
                return (int)longValue;
                
            if (int.TryParse(value.ToString(), out int parsedValue))
                return parsedValue;
                
            return defaultValue;
        }

        static async Task<List<Repository>> GetAllUserRepositories(string username)
        {
            var repositories = new List<Repository>();
            int page = 1;

            while (page <= 10) // 最多10頁，獲取更多倉庫
            {
                var url = $"https://api.github.com/users/{username}/repos?page={page}&per_page=100&sort=stars&direction=desc";
                var response = await MakeGitHubApiCall<List<dynamic>>(url);
                
                if (!response.IsSuccess)
                    break;

                if (response.Data.Count == 0)
                    break;

                foreach (var repo in response.Data)
                {
                    var starCount = repo.stargazers_count ?? 0;
                    var forkCount = repo.forks_count ?? 0;
                    var isFork = SafeGetBool(repo.fork);
                    var ownerLogin = repo.owner?.login?.ToString() ?? "";
                    
                    // 如果是 Fork 專案，檢查原始專案的貢獻者排名
                    if (isFork && repo.parent != null)
                    {
                        var parentFullName = repo.parent.full_name?.ToString();
                        var parentOwnerType = repo.parent.owner?.type?.ToString() ?? "";
                        var parentOwnerLogin = repo.parent.owner?.login?.ToString() ?? "";
                        
                        if (!string.IsNullOrEmpty(parentFullName))
                        {
                            // 檢查用戶是否為原始專案的前10名貢獻者
                            var isTopContributor = await IsUserTopContributor(username, parentFullName, 10);
                            
                            if (isTopContributor)
                            {
                                var parentStarCount = repo.parent.stargazers_count ?? 0;
                                var parentForkCount = repo.parent.forks_count ?? 0;
                                
                                // 只保存兩顆星以上的原始專案
                                if (parentStarCount >= 2)
                                {
                                    repositories.Add(new Repository
                                    {
                                        Name = repo.parent.name?.ToString() ?? "",
                                        FullName = parentFullName ?? "",
                                        StargazersCount = parentStarCount,
                                        ForksCount = parentForkCount,
                                        HtmlUrl = repo.parent.html_url?.ToString() ?? "",
                                        Language = repo.parent.language?.ToString() ?? "",
                                        IsFork = false, // 原始專案本身不是 Fork
                                        OwnerLogin = parentOwnerLogin,
                                        IsOrganization = parentOwnerType == "Organization"
                                    });
                                    
                                    Console.WriteLine($"發現 {username} 是 {parentFullName} 的前10名貢獻者 (通過 Fork 專案 {repo.full_name} 發現)");
                                }
                            }
                        }
                    }
                    else if (!isFork)
                    {
                        // 處理非 Fork 的個人專案
                        // 只保存兩顆星以上的專案
                        if (starCount >= 2)
                        {
                            repositories.Add(new Repository
                            {
                                Name = repo.name,
                                FullName = repo.full_name,
                                StargazersCount = starCount,
                                ForksCount = forkCount,
                                HtmlUrl = repo.html_url ?? "",
                                Language = repo.language ?? "",
                                IsFork = false,
                                OwnerLogin = ownerLogin,
                                IsOrganization = false
                            });
                        }
                    }
                }

                page++;
                await Task.Delay(100);
            }

            return repositories;
        }

        static async Task<List<Repository>> GetUserRepositories(string username)
        {
            var repositories = new List<Repository>();
            int page = 1;

            while (page <= 5) // 最多5頁
            {
                var url = $"https://api.github.com/users/{username}/repos?page={page}&per_page=100&sort=stars&direction=desc";
                var response = await MakeGitHubApiCall<List<dynamic>>(url);
                
                if (!response.IsSuccess)
                    break;

                if (response.Data.Count == 0)
                    break;

                foreach (var repo in response.Data)
                {
                    var starCount = repo.stargazers_count ?? 0;
                    var forkCount = repo.forks_count ?? 0;
                    var isFork = SafeGetBool(repo.fork);
                    
                    // 只保存兩顆星以上的專案
                    if (starCount >= 2)
                    {
                        repositories.Add(new Repository
                        {
                            Name = repo.name?.ToString() ?? "",
                            FullName = repo.full_name?.ToString() ?? "",
                            StargazersCount = starCount,
                            ForksCount = forkCount,
                            HtmlUrl = repo.html_url?.ToString() ?? "",
                            Language = repo.language?.ToString() ?? "",
                            IsFork = isFork,
                            OwnerLogin = repo.owner?.login?.ToString() ?? "",
                            IsOrganization = false
                        });
                    }
                }

                page++;
                await Task.Delay(100);
            }

            return repositories;
        }

        static async Task<List<Repository>> GetUserOrganizationRepositories(string username)
        {
            var repositories = new List<Repository>();
            
            // 獲取用戶參與的組織
            var orgsUrl = $"https://api.github.com/users/{username}/orgs";
            var orgsResponse = await MakeGitHubApiCall<List<dynamic>>(orgsUrl);
            
            if (!orgsResponse.IsSuccess)
                return repositories;

            foreach (var org in orgsResponse.Data)
            {
                var orgLogin = org.login;
                
                // 獲取組織的倉庫
                var orgReposUrl = $"https://api.github.com/orgs/{orgLogin}/repos?sort=stars&direction=desc&per_page=100";
                var orgReposResponse = await MakeGitHubApiCall<List<dynamic>>(orgReposUrl);
                
                if (orgReposResponse.IsSuccess)
                {
                    foreach (var repo in orgReposResponse.Data.Take(20)) // 每個組織最多20個倉庫
                    {
                        var repoFullName = repo.full_name?.ToString();
                        if (string.IsNullOrEmpty(repoFullName)) continue;
                        
                        // 獲取貢獻者排名信息
                        var rankInfo = await GetContributorRankInfo(username, repoFullName);
                        
                        if (rankInfo.IsContributor)
                        {
                            var starCount = repo.stargazers_count ?? 0;
                            var forkCount = repo.forks_count ?? 0;
                            var isFork = SafeGetBool(repo.fork);
                            
                            // 只保存兩顆星以上的專案
                            if (starCount >= 2)
                            {
                                repositories.Add(new Repository
                                {
                                    Name = repo.name?.ToString() ?? "",
                                    FullName = repoFullName ?? "",
                                    StargazersCount = starCount,
                                    ForksCount = forkCount,
                                    HtmlUrl = repo.html_url?.ToString() ?? "",
                                    Language = repo.language?.ToString() ?? "",
                                    IsFork = isFork,
                                    OwnerLogin = orgLogin?.ToString() ?? "",
                                    IsOrganization = true,
                                    ContributorRank = rankInfo.Rank,
                                    TotalContributors = rankInfo.TotalContributors
                                });
                                
                                Console.WriteLine($"發現 {username} 在組織專案 {repoFullName} 的貢獻排名: {rankInfo.Rank}/{rankInfo.TotalContributors} ({rankInfo.CommitCount} commits)");
                            }
                        }
                        
                        await Task.Delay(50);
                    }
                }
                
                await Task.Delay(100);
            }

            return repositories;
        }

        static async Task<List<Repository>> GetUserContributedRepositories(string username)
        {
            var repositories = new List<Repository>();
            var processedRepos = new HashSet<string>();
            
            try
            {
                // 使用 GitHub Search Issues API 來尋找用戶提交的 PR
                // 搜尋該用戶作為 PR 作者的專案，但排除自己擁有的專案
                var searchQuery = $"is:pr+author:{username}";
                var searchUrl = $"https://api.github.com/search/issues?q={searchQuery}&sort=updated&order=desc&per_page=100";
                
                var searchResponse = await MakeGitHubApiCall<dynamic>(searchUrl);
                if (!searchResponse.IsSuccess)
                {
                    Console.WriteLine($"搜尋 PR 貢獻專案時發生錯誤: {searchResponse.ErrorMessage}");
                    return repositories;
                }

                var items = searchResponse.Data?.items;
                if (items == null)
                    return repositories;

                Console.WriteLine($"找到 {items.Count} 個 PR，正在分析相關專案...");

                foreach (var pr in items)
                {
                    var repoUrl = pr.repository_url?.ToString();
                    if (string.IsNullOrEmpty(repoUrl)) continue;
                    
                    // 從 repository_url 取得專案詳細資訊
                    var repoResponse = await MakeGitHubApiCall<dynamic>(repoUrl);
                    if (!repoResponse.IsSuccess) continue;
                    
                    var repo = repoResponse.Data;
                    var ownerType = repo.owner?.type?.ToString() ?? "";
                    var ownerLogin = repo.owner?.login?.ToString() ?? "";
                    var repoFullName = repo.full_name?.ToString() ?? "";
                    
                    // 只取個人專案（非組織專案），且不是自己的專案
                    if (ownerType == "User" && ownerLogin != username && !processedRepos.Contains(repoFullName))
                    {
                        processedRepos.Add(repoFullName);
                        
                        // 獲取貢獻者排名信息
                        var rankInfo = await GetContributorRankInfo(username, repoFullName);
                        
                        if (rankInfo.IsContributor)
                        {
                            var starCount = repo.stargazers_count ?? 0;
                            var forkCount = repo.forks_count ?? 0;
                            
                            // 只保存兩顆星以上的專案
                            if (starCount >= 2)
                            {
                                var isFork = SafeGetBool(repo.fork);
                                
                                repositories.Add(new Repository
                                {
                                    Name = repo.name?.ToString() ?? "",
                                    FullName = repoFullName,
                                    StargazersCount = starCount,
                                    ForksCount = forkCount,
                                    HtmlUrl = repo.html_url?.ToString() ?? "",
                                    Language = repo.language?.ToString() ?? "",
                                    IsFork = isFork,
                                    OwnerLogin = ownerLogin,
                                    IsOrganization = false,
                                    ContributorRank = rankInfo.Rank,
                                    TotalContributors = rankInfo.TotalContributors
                                });
                                
                                Console.WriteLine($"發現 {username} 在個人專案 {repoFullName} 的貢獻排名: {rankInfo.Rank}/{rankInfo.TotalContributors} ({rankInfo.CommitCount} commits, ⭐{starCount})");
                            }
                        }
                        
                        await Task.Delay(100); // 避免API限制
                    }
                    
                    // 限制最多檢查前30個專案以避免API限制
                    if (repositories.Count >= 30)
                        break;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"獲取用戶 {username} 的貢獻專案時發生異常: {ex.Message}");
            }

            return repositories;
        }

        /// <summary>
        /// 從URL提取API端點名稱用於統計，保留具體參數值
        /// </summary>
        /// <param name="url">完整的API URL</param>
        /// <returns>帶具體參數的API端點名稱</returns>
        static string ExtractApiEndpoint(string url)
        {
            try
            {
                var uri = new Uri(url);
                var path = uri.AbsolutePath;
                
                // 移除開頭的 /
                if (path.StartsWith("/"))
                    path = path.Substring(1);
                
                // 為了便於分析，保留具體的參數值
                if (path.StartsWith("search/"))
                {
                    // 對於搜索API，顯示查詢參數
                    var queryString = uri.Query;
                    if (!string.IsNullOrEmpty(queryString))
                    {
                        // 手動解析查詢參數來獲取q參數
                        var queryPairs = queryString.TrimStart('?').Split('&');
                        foreach (var pair in queryPairs)
                        {
                            var keyValue = pair.Split('=');
                            if (keyValue.Length == 2 && keyValue[0] == "q")
                            {
                                var query = Uri.UnescapeDataString(keyValue[1]);
                                // 只顯示前50個字符的查詢內容
                                var displayQuery = query.Length > 50 ? query.Substring(0, 50) + "..." : query;
                                return $"search/{path.Split('/')[1]}?q={displayQuery}";
                            }
                        }
                    }
                    return "search/" + path.Split('/')[1];
                }
                else if (path.StartsWith("users/"))
                {
                    var parts = path.Split('/');
                    if (parts.Length == 2)
                        return $"users/{parts[1]}"; // users/{username}
                    else if (parts.Length >= 3)
                        return $"users/{parts[1]}/" + string.Join("/", parts.Skip(2)); // users/{username}/repos
                }
                else if (path.StartsWith("repos/"))
                {
                    var parts = path.Split('/');
                    if (parts.Length >= 3)
                    {
                        var owner = parts[1];
                        var repo = parts[2];
                        var endpoint = parts.Length > 3 ? "/" + string.Join("/", parts.Skip(3)) : "";
                        return $"repos/{owner}/{repo}{endpoint}";
                    }
                }
                else if (path.StartsWith("orgs/"))
                {
                    var parts = path.Split('/');
                    if (parts.Length >= 2)
                    {
                        var org = parts[1];
                        var endpoint = parts.Length > 2 ? "/" + string.Join("/", parts.Skip(2)) : "";
                        return $"orgs/{org}{endpoint}";
                    }
                }
                
                return path;
            }
            catch
            {
                return url; // 如果解析失敗，返回原始URL
            }
        }

        /// <summary>
        /// 記錄API調用統計
        /// </summary>
        /// <param name="endpoint">API端點</param>
        /// <param name="elapsedMs">耗時（毫秒）</param>
        static void RecordApiCall(string endpoint, long elapsedMs)
        {
            lock (statsLock)
            {
                if (!apiCallCounts.ContainsKey(endpoint))
                {
                    apiCallCounts[endpoint] = 0;
                    apiCallTimes[endpoint] = 0;
                }
                
                apiCallCounts[endpoint]++;
                apiCallTimes[endpoint] += elapsedMs;
                
                Console.WriteLine($"[API統計] {endpoint}: 第{apiCallCounts[endpoint]}次調用, 耗時{elapsedMs}ms, 累計{apiCallTimes[endpoint]}ms");
            }
        }

        /// <summary>
        /// 記錄錯誤統計
        /// </summary>
        /// <param name="errorType">錯誤類型</param>
        static void RecordError(string errorType)
        {
            lock (statsLock)
            {
                if (!errorCounts.ContainsKey(errorType))
                {
                    errorCounts[errorType] = 0;
                }
                errorCounts[errorType]++;
            }
        }

        /// <summary>
        /// 清理過期的緩存
        /// </summary>
        static void CleanExpiredCache()
        {
            lock (apiCacheLock)
            {
                var expiredKeys = apiCacheTimestamps
                    .Where(kvp => DateTime.Now - kvp.Value > cacheExpireTime)
                    .Select(kvp => kvp.Key)
                    .ToList();
                
                foreach (var key in expiredKeys)
                {
                    apiResponseCache.Remove(key);
                    apiCacheTimestamps.Remove(key);
                }
                
                if (expiredKeys.Any())
                {
                    Console.WriteLine($"清理了 {expiredKeys.Count} 個過期緩存條目");
                }
            }
        }

        /// <summary>
        /// 打印緩存統計信息
        /// </summary>
        static void PrintCacheStats()
        {
            lock (apiCacheLock)
            {
                Console.WriteLine("\n=== API緩存統計 ===");
                Console.WriteLine($"緩存條目總數: {apiResponseCache.Count}");
                
                var expiredCount = 0;
                var validCount = 0;
                
                foreach (var timestamp in apiCacheTimestamps.Values)
                {
                    if (DateTime.Now - timestamp > cacheExpireTime)
                        expiredCount++;
                    else
                        validCount++;
                }
                
                Console.WriteLine($"有效緩存: {validCount}");
                Console.WriteLine($"過期緩存: {expiredCount}");
                Console.WriteLine($"緩存過期時間: {cacheExpireTime.TotalMinutes} 分鐘");
                Console.WriteLine("==================\n");
            }
        }

        /// <summary>
        /// 打印API調用總結統計
        /// </summary>
        static void PrintApiStatsSummary()
        {
            lock (statsLock)
            {
                Console.WriteLine("\n=== API調用統計總結 ===");
                Console.WriteLine($"{"API端點",-30} {"調用次數",10} {"總耗時(ms)",12} {"平均耗時(ms)",12}");
                Console.WriteLine(new string('-', 70));
                
                var totalCalls = 0;
                var totalTime = 0L;
                
                foreach (var endpoint in apiCallCounts.Keys.OrderBy(k => k))
                {
                    var count = apiCallCounts[endpoint];
                    var time = apiCallTimes[endpoint];
                    var avgTime = count > 0 ? time / count : 0;
                    
                    Console.WriteLine($"{endpoint,-30} {count,10} {time,12} {avgTime,12}");
                    totalCalls += count;
                    totalTime += time;
                }
                
                Console.WriteLine(new string('-', 70));
                var totalAvgTime = totalCalls > 0 ? totalTime / totalCalls : 0;
                Console.WriteLine($"{"總計",-30} {totalCalls,10} {totalTime,12} {totalAvgTime,12}");
                
                // 添加錯誤統計
                if (errorCounts.Any() || totalRetries > 0 || longWaits > 0)
                {
                    Console.WriteLine();
                    Console.WriteLine("=== 錯誤統計 ===");
                    Console.WriteLine($"總重試次數: {totalRetries}");
                    Console.WriteLine($"長時間等待次數: {longWaits}");
                    
                    if (errorCounts.Any())
                    {
                        Console.WriteLine("錯誤類型統計:");
                        foreach (var error in errorCounts.OrderByDescending(e => e.Value))
                        {
                            Console.WriteLine($"  {error.Key}: {error.Value} 次");
                        }
                    }
                }
                
                Console.WriteLine("========================\n");
            }
        }

        static async Task<GitHubApiResponse<T>> MakeGitHubApiCall<T>(string url)
        {
            var cacheKey = GenerateCacheKey(url);
            var shouldCache = ShouldCacheUrl(url);
            
            // 首先檢查緩存（只有當URL應該被緩存時才檢查）
            if (shouldCache)
            {
                var cachedResponse = GetFromCache<T>(cacheKey);
                if (cachedResponse != null)
                {
                    return cachedResponse;
                }
            }
            
            var endpoint = ExtractApiEndpoint(url);
            var stopwatch = System.Diagnostics.Stopwatch.StartNew();
            
            int retryCount = 0;
            const int maxRetries = 3;
            int longWaitCount = 0;

            while (true) // 無限重試，直到成功或遇到不可恢復的錯誤
            {
                try
                {
                    var response = await httpClient.GetAsync(url);
                    var content = await response.Content.ReadAsStringAsync();

                    if (response.IsSuccessStatusCode)
                    {
                        // 預先檢查是否為空物件或期待列表但收到物件的情況
                        var trimmedContent = content.Trim();
                        
                        // 處理 GitHub stats/contributors API 返回 {} 的情況
                        if (trimmedContent == "{}" || trimmedContent == "[]")
                        {
                            // 靜默處理，不輸出錯誤訊息，因為這是 GitHub API 的正常行為
                            if (typeof(T).IsGenericType && typeof(T).GetGenericTypeDefinition() == typeof(List<>))
                            {
                                var emptyList = Activator.CreateInstance<T>();
                                var emptyListResponse = new GitHubApiResponse<T>
                                {
                                    Data = emptyList,
                                    IsSuccess = true,
                                    RemainingRequests = int.Parse(response.Headers.GetValues("X-RateLimit-Remaining").FirstOrDefault() ?? "0"),
                                    ResetTime = DateTimeOffset.FromUnixTimeSeconds(long.Parse(response.Headers.GetValues("X-RateLimit-Reset").FirstOrDefault() ?? "0")).DateTime
                                };
                                
                                // 保存到緩存
                                if (shouldCache)
                                {
                                    SaveToCache(cacheKey, emptyListResponse);
                                }
                                
                                return emptyListResponse;
                            }
                        }
                        
                        // 檢查是否期待 List 但收到單一物件
                        if (typeof(T).IsGenericType && typeof(T).GetGenericTypeDefinition() == typeof(List<>) && 
                            trimmedContent.StartsWith("{") && trimmedContent.EndsWith("}") && !trimmedContent.StartsWith("["))
                        {
                            // 期待陣列但收到物件，靜默返回空列表
                            var emptyList = Activator.CreateInstance<T>();
                            var emptyListResponse = new GitHubApiResponse<T>
                            {
                                Data = emptyList,
                                IsSuccess = true,
                                RemainingRequests = int.Parse(response.Headers.GetValues("X-RateLimit-Remaining").FirstOrDefault() ?? "0"),
                                ResetTime = DateTimeOffset.FromUnixTimeSeconds(long.Parse(response.Headers.GetValues("X-RateLimit-Reset").FirstOrDefault() ?? "0")).DateTime
                            };
                            
                            // 保存到緩存
                            if (shouldCache)
                            {
                                SaveToCache(cacheKey, emptyListResponse);
                            }
                            
                            return emptyListResponse;
                        }
                        
                        try
                        {
                            var data = JsonConvert.DeserializeObject<T>(content);
                            stopwatch.Stop();
                            RecordApiCall(endpoint, stopwatch.ElapsedMilliseconds);
                            
                            var apiResponse = new GitHubApiResponse<T>
                            {
                                Data = data!,
                                IsSuccess = true,
                                RemainingRequests = int.Parse(response.Headers.GetValues("X-RateLimit-Remaining").FirstOrDefault() ?? "0"),
                                ResetTime = DateTimeOffset.FromUnixTimeSeconds(long.Parse(response.Headers.GetValues("X-RateLimit-Reset").FirstOrDefault() ?? "0")).DateTime
                            };
                            
                            // 保存到緩存 (只緩存成功的響應且URL應該被緩存)
                            if (shouldCache)
                            {
                                SaveToCache(cacheKey, apiResponse);
                            }
                            
                            return apiResponse;
                        }
                        catch (JsonSerializationException ex)
                        {
                            // 如果預檢查沒有捕獲到，但仍然發生反序列化錯誤
                            Console.WriteLine($"JSON 反序列化錯誤，URL: {url}");
                            Console.WriteLine($"回應內容: {content.Substring(0, Math.Min(200, content.Length))}...");
                            Console.WriteLine($"錯誤詳情: {ex.Message}");
                            
                            // 最後的回退處理
                            if (typeof(T).IsGenericType && typeof(T).GetGenericTypeDefinition() == typeof(List<>))
                            {
                                Console.WriteLine("回退到空列表處理");
                                var emptyList = Activator.CreateInstance<T>();
                                stopwatch.Stop();
                                RecordApiCall(endpoint, stopwatch.ElapsedMilliseconds);
                                
                                var emptyListResponse = new GitHubApiResponse<T>
                                {
                                    Data = emptyList,
                                    IsSuccess = true,
                                    RemainingRequests = int.Parse(response.Headers.GetValues("X-RateLimit-Remaining").FirstOrDefault() ?? "0"),
                                    ResetTime = DateTimeOffset.FromUnixTimeSeconds(long.Parse(response.Headers.GetValues("X-RateLimit-Reset").FirstOrDefault() ?? "0")).DateTime
                                };
                                
                                // 保存到緩存
                                if (shouldCache)
                                {
                                    SaveToCache(cacheKey, emptyListResponse);
                                }
                                
                                return emptyListResponse;
                            }
                            
                            stopwatch.Stop();
                            RecordApiCall(endpoint, stopwatch.ElapsedMilliseconds);
                            
                            return new GitHubApiResponse<T>
                            {
                                IsSuccess = false,
                                ErrorMessage = $"JSON 反序列化失敗: {ex.Message}"
                            };
                        }
                    }
                    else if (response.StatusCode == System.Net.HttpStatusCode.Accepted)
                    {
                        // 202 Accepted - GitHub stats API 正在計算資料，返回空結果
                        Console.WriteLine($"GitHub API 正在計算統計資料，稍後重試: {url}");
                        
                        // 對於 List 類型，返回空列表
                        if (typeof(T).IsGenericType && typeof(T).GetGenericTypeDefinition() == typeof(List<>))
                        {
                            var emptyList = Activator.CreateInstance<T>();
                            var acceptedResponse = new GitHubApiResponse<T>
                            {
                                Data = emptyList,
                                IsSuccess = true,
                                RemainingRequests = int.Parse(response.Headers.GetValues("X-RateLimit-Remaining").FirstOrDefault() ?? "0"),
                                ResetTime = DateTimeOffset.FromUnixTimeSeconds(long.Parse(response.Headers.GetValues("X-RateLimit-Reset").FirstOrDefault() ?? "0")).DateTime
                            };
                            
                            // 不緩存 Accepted 響應，因為數據可能稍後可用
                            
                            return acceptedResponse;
                        }
                        
                        return new GitHubApiResponse<T>
                        {
                            IsSuccess = false,
                            ErrorMessage = "API 正在計算資料"
                        };
                    }
                    else if (response.StatusCode == System.Net.HttpStatusCode.Unauthorized)
                    {
                        // 401 Unauthorized - Token 問題
                        var errorMsg = "GitHub API Token 驗證失敗。請檢查：\n" +
                                      "1. Token 是否正確（應以 ghp_ 或 github_pat_ 開頭）\n" +
                                      "2. Token 是否已過期\n" +
                                      "3. Token 是否有適當的權限（至少需要 public_repo 權限）\n" +
                                      "4. 請到 https://github.com/settings/tokens 檢查或重新產生 Token";
                        
                        return new GitHubApiResponse<T>
                        {
                            IsSuccess = false,
                            ErrorMessage = errorMsg
                        };
                    }
                    else if (response.StatusCode == System.Net.HttpStatusCode.Forbidden)
                    {
                        Console.WriteLine($"GitHub API 403 Forbidden 錯誤:");
                        Console.WriteLine($"  API端點: {endpoint}");
                        Console.WriteLine($"  回傳內容: {content}");
                        
                        // 檢查是否是貢獻者列表過大的特殊錯誤
                        if (content.Contains("too large to list contributors") || 
                            content.Contains("contributor list is too large"))
                        {
                            Console.WriteLine($"  錯誤原因: 貢獻者列表過大，GitHub拒絕提供數據");
                            Console.WriteLine($"  建議: 此專案貢獻者過多，跳過排名檢查");
                            return new GitHubApiResponse<T>
                            {
                                IsSuccess = false,
                                ErrorMessage = content
                            };
                        }
                        
                        // API限制，等待後重試
                        var remainingRequests = 0;
                        var resetTime = DateTime.UtcNow;
                        
                        try
                        {
                            remainingRequests = int.Parse(response.Headers.GetValues("X-RateLimit-Remaining").FirstOrDefault() ?? "0");
                            resetTime = DateTimeOffset.FromUnixTimeSeconds(long.Parse(response.Headers.GetValues("X-RateLimit-Reset").FirstOrDefault() ?? "0")).DateTime;
                        }
                        catch
                        {
                            Console.WriteLine($"  警告: 無法解析速率限制標頭");
                        }
                        
                        var waitTime = resetTime - DateTime.UtcNow;
                        Console.WriteLine($"  剩餘請求數: {remainingRequests}");
                        Console.WriteLine($"  重置時間: {resetTime:yyyy-MM-dd HH:mm:ss} UTC");
                        Console.WriteLine($"  等待時間: {waitTime.TotalMinutes:F1} 分鐘");
                        
                        if (waitTime.TotalSeconds > 0)
                        {
                            Console.WriteLine($"GitHub API限制，等待 {waitTime.TotalMinutes:F1} 分鐘後重試...");
                            await Task.Delay((int)waitTime.TotalMilliseconds + 10000);
                        }
                        else
                        {
                            Console.WriteLine($"無法確定等待時間，等待5分鐘後重試...");
                            await Task.Delay(300000); // 等待5分鐘
                        }
                        
                        retryCount++;
                        continue;
                    }
                    else if (response.StatusCode == System.Net.HttpStatusCode.ServiceUnavailable || 
                             response.StatusCode == System.Net.HttpStatusCode.BadGateway ||
                             response.StatusCode == System.Net.HttpStatusCode.InternalServerError)
                    {
                        // 記錄錯誤統計
                        RecordError($"服務不可用_{response.StatusCode}");
                        totalRetries++;
                        
                        // 服務不可用錯誤，重試
                        retryCount++;
                        Console.WriteLine($"GitHub API 服務不可用 ({response.StatusCode})，第 {retryCount} 次重試...");
                        Console.WriteLine($"  API端點: {endpoint}");
                        
                        // 如果回傳內容為空，顯示更友好的信息
                        if (string.IsNullOrWhiteSpace(content))
                        {
                            Console.WriteLine($"  回傳內容: [空回應] - GitHub伺服器可能正在維護或過載");
                            Console.WriteLine($"  建議: 這是GitHub伺服器的臨時問題，通常重新運行程序即可解決");
                        }
                        else
                        {
                            Console.WriteLine($"  回傳內容: {(content.Length > 200 ? content.Substring(0, 200) + "..." : content)}");
                        }
                        
                        // 對於stats/contributors API的特殊處理
                        if (endpoint.Contains("/stats/contributors"))
                        {
                            Console.WriteLine($"  特殊情況: GitHub統計API可能正在重新計算數據");
                            Console.WriteLine($"  說明: stats/contributors API在大型專案上經常出現臨時錯誤");
                            
                            if (retryCount >= maxRetries)
                            {
                                Console.WriteLine($"  跳過此統計API，改用基本contributors API作為替代方案");
                                // 對於stats API，我們可以返回空結果而不是無限重試
                                if (typeof(T).IsGenericType && typeof(T).GetGenericTypeDefinition() == typeof(List<>))
                                {
                                    var emptyList = Activator.CreateInstance<T>();
                                    stopwatch.Stop();
                                    RecordApiCall(endpoint, stopwatch.ElapsedMilliseconds);
                                    
                                    return new GitHubApiResponse<T>
                                    {
                                        Data = emptyList,
                                        IsSuccess = true,
                                        RemainingRequests = 0,
                                        ResetTime = DateTime.UtcNow.AddHours(1)
                                    };
                                }
                            }
                        }
                        
                        if (retryCount >= maxRetries)
                        {
                            Console.WriteLine("GitHub API 服務持續不可用，重新創建HttpClient並等待10分鐘後繼續重試...");
                            Console.WriteLine($"  已重試 {maxRetries} 次，錯誤狀態: {response.StatusCode}");
                            Console.WriteLine($"  下次長等待將是第 {longWaitCount + 1} 次");
                            if (!string.IsNullOrWhiteSpace(content))
                            {
                                Console.WriteLine($"  完整回傳內容: {content}");
                            }
                            
                            // 重新創建HttpClient
                            RecreateHttpClient();
                            
                            longWaitCount++;
                            longWaits++; // 全局統計
                            retryCount = 0; // 重置重試計數器
                            await Task.Delay(600000); // 等待10分鐘
                            Console.WriteLine($"第 {longWaitCount} 次長時間等待完成，重新開始重試...");
                            continue;
                        }
                        
                        // 對於空回應，使用較短的等待時間
                        var waitTime = string.IsNullOrWhiteSpace(content) ? 2000 : 5000;
                        await Task.Delay(waitTime * retryCount); // 遞增等待時間
                        continue;
                    }
                    else
                    {
                        // 檢查是否是服務不可用的錯誤訊息
                        if (content.Contains("No server is currently available") || 
                            content.Contains("service your request") ||
                            string.IsNullOrWhiteSpace(content)) // 新增：空回應也視為服務不可用
                        {
                            retryCount++;
                            Console.WriteLine($"GitHub API 服務不可用，第 {retryCount} 次重試...");
                            Console.WriteLine($"  API端點: {endpoint}");
                            Console.WriteLine($"  HTTP狀態碼: {response.StatusCode}");
                            
                            if (string.IsNullOrWhiteSpace(content))
                            {
                                Console.WriteLine($"  錯誤訊息: [空回應] - 可能是GitHub伺服器臨時問題");
                                Console.WriteLine($"  常見原因: 伺服器過載、維護、或網路連線不穩定");
                            }
                            else
                            {
                                Console.WriteLine($"  錯誤訊息: {(content.Length > 200 ? content.Substring(0, 200) + "..." : content)}");
                            }
                            
                            if (retryCount >= maxRetries)
                            {
                                Console.WriteLine("GitHub API 服務持續不可用，重新創建HttpClient並等待10分鐘後繼續重試...");
                                Console.WriteLine($"  已重試 {maxRetries} 次，錯誤狀態: {response.StatusCode}");
                                Console.WriteLine($"  建議: 如果問題持續，可以:");
                                Console.WriteLine($"    1. 檢查GitHub狀態頁面: https://status.github.com/");
                                Console.WriteLine($"    2. 稍後重新運行程序");
                                Console.WriteLine($"    3. 檢查網路連線是否穩定");
                                if (!string.IsNullOrWhiteSpace(content))
                                {
                                    Console.WriteLine($"  完整錯誤內容: {content}");
                                }
                                
                                // 重新創建HttpClient
                                RecreateHttpClient();
                                
                                longWaitCount++;
                                retryCount = 0; // 重置重試計數器
                                await Task.Delay(600000); // 等待10分鐘
                                Console.WriteLine($"第 {longWaitCount} 次長時間等待完成，重新開始重試...");
                                continue;
                            }
                            
                            // 對於空回應，使用較短的等待時間
                            var waitTime = string.IsNullOrWhiteSpace(content) ? 2000 : 5000;
                            await Task.Delay(waitTime * retryCount); // 遞增等待時間
                            continue;
                        }
                        
                        // 其他不可恢復的錯誤
                        stopwatch.Stop();
                        RecordApiCall(endpoint, stopwatch.ElapsedMilliseconds);
                        Console.WriteLine($"不可恢復的API錯誤:");
                        Console.WriteLine($"  API端點: {endpoint}");
                        Console.WriteLine($"  HTTP狀態碼: {response.StatusCode}");
                        Console.WriteLine($"  回傳內容: {content}");
                        Console.WriteLine($"  調用耗時: {stopwatch.ElapsedMilliseconds}ms");
                        
                        return new GitHubApiResponse<T>
                        {
                            IsSuccess = false,
                            ErrorMessage = $"HTTP {response.StatusCode}: {content}"
                        };
                    }
                }
                catch (HttpRequestException ex)
                {
                    retryCount++;
                    Console.WriteLine($"網路連線問題: {ex.Message}，第 {retryCount} 次重試...");
                    Console.WriteLine($"  API端點: {endpoint}");
                    Console.WriteLine($"  異常類型: HttpRequestException");
                    Console.WriteLine($"  詳細信息: {ex}");
                    
                    if (retryCount >= maxRetries)
                    {
                        Console.WriteLine("網路連線持續有問題，重新創建HttpClient並等待10分鐘後繼續重試...");
                        Console.WriteLine($"  已重試 {maxRetries} 次，網路連線異常");
                        Console.WriteLine($"  原因: 無法建立HTTP連線到GitHub API");
                        
                        // 重新創建HttpClient
                        RecreateHttpClient();
                        
                        longWaitCount++;
                        retryCount = 0; // 重置重試計數器
                        await Task.Delay(600000); // 等待10分鐘
                        Console.WriteLine($"第 {longWaitCount} 次長時間等待完成，重新開始重試...");
                        continue;
                    }
                    
                    await Task.Delay(5000 * retryCount); // 遞增等待時間
                }
                catch (TaskCanceledException ex)
                {
                    retryCount++;
                    Console.WriteLine($"請求超時: {ex.Message}，第 {retryCount} 次重試...");
                    Console.WriteLine($"  API端點: {endpoint}");
                    Console.WriteLine($"  異常類型: TaskCanceledException");
                    Console.WriteLine($"  超時時間: {httpClient.Timeout.TotalSeconds} 秒");
                    
                    if (retryCount >= maxRetries)
                    {
                        Console.WriteLine("請求持續超時，重新創建HttpClient並等待10分鐘後繼續重試...");
                        Console.WriteLine($"  已重試 {maxRetries} 次，請求超時");
                        Console.WriteLine($"  原因: HTTP請求超過設定的超時時間");
                        
                        // 重新創建HttpClient
                        RecreateHttpClient();
                        
                        longWaitCount++;
                        retryCount = 0; // 重置重試計數器
                        await Task.Delay(600000); // 等待10分鐘
                        Console.WriteLine($"第 {longWaitCount} 次長時間等待完成，重新開始重試...");
                        continue;
                    }
                    
                    await Task.Delay(5000 * retryCount); // 遞增等待時間
                }
                catch (Exception ex)
                {
                    retryCount++;
                    Console.WriteLine($"API 調用異常: {ex.Message}，第 {retryCount} 次重試...");
                    Console.WriteLine($"  API端點: {endpoint}");
                    Console.WriteLine($"  異常類型: {ex.GetType().Name}");
                    Console.WriteLine($"  堆疊追蹤: {ex.StackTrace}");
                    
                    if (retryCount >= maxRetries)
                    {
                        stopwatch.Stop();
                        RecordApiCall(endpoint, stopwatch.ElapsedMilliseconds);
                        
                        Console.WriteLine($"不可恢復的API調用異常:");
                        Console.WriteLine($"  API端點: {endpoint}");
                        Console.WriteLine($"  已重試 {maxRetries} 次");
                        Console.WriteLine($"  最終異常: {ex.GetType().Name} - {ex.Message}");
                        Console.WriteLine($"  調用耗時: {stopwatch.ElapsedMilliseconds}ms");
                        Console.WriteLine($"  完整異常信息: {ex}");
                        
                        return new GitHubApiResponse<T>
                        {
                            IsSuccess = false,
                            ErrorMessage = ex.Message
                        };
                    }
                    
                    await Task.Delay(1000 * retryCount); // 指數退避
                }
            }
        }

        static async Task<bool> ValidateGitHubToken()
        {
            try
            {
                // 直接使用 httpClient 進行驗證，不使用 MakeGitHubApiCall 避免重複錯誤處理
                var response = await httpClient.GetAsync("https://api.github.com/user");
                
                if (response.IsSuccessStatusCode)
                {
                    return true;
                }
                else if (response.StatusCode == System.Net.HttpStatusCode.Unauthorized)
                {
                    var content = await response.Content.ReadAsStringAsync();
                    Console.WriteLine($"Token 驗證失敗: {content}");
                    return false;
                }
                else
                {
                    Console.WriteLine($"Token 驗證遇到問題: {response.StatusCode}");
                    return false;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Token 驗證時發生異常: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 設置HttpClient的Headers和Authorization
        /// </summary>
        static void SetupHttpClient()
        {
            try
            {
                // 清理Headers
                httpClient.DefaultRequestHeaders.Clear();
                
                // 設置基本Headers
                httpClient.DefaultRequestHeaders.UserAgent.ParseAdd("TaiwanPopularDevelopers/1.0");
                httpClient.DefaultRequestHeaders.Add("Accept", "application/vnd.github.v3+json");
                
                // 設置超時時間
                httpClient.Timeout = TimeSpan.FromMinutes(5);
                
                // 如果有Token，設置Authorization
                if (!string.IsNullOrEmpty(githubToken))
                {
                    httpClient.DefaultRequestHeaders.Authorization = 
                        new System.Net.Http.Headers.AuthenticationHeaderValue("token", githubToken);
                }
                
                Console.WriteLine("HttpClient 設置完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"設置HttpClient時發生錯誤: {ex.Message}");
            }
        }

        /// <summary>
        /// 重新創建HttpClient實例
        /// </summary>
        static void RecreateHttpClient()
        {
            try
            {
                // 清理舊的HttpClient
                httpClient?.Dispose();
                
                // 創建新的HttpClient
                httpClient = new HttpClient();
                
                // 重新設置HttpClient
                SetupHttpClient();
                
                Console.WriteLine("已重新創建HttpClient實例");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"重新創建HttpClient時發生錯誤: {ex.Message}");
            }
        }

        static void ClearHttpClientHeaders()
        {
            try
            {
                httpClient.DefaultRequestHeaders.Clear();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"清理 HTTP headers 時發生異常: {ex.Message}");
            }
        }


        static string GenerateMarkdown(List<GitHubUser> users)
        {
            var sb = new StringBuilder();
            
            sb.AppendLine($"# {currentRegion.ChineseName}知名GitHub用戶排名");
            sb.AppendLine();
            sb.AppendLine("> 本排名基於以下指標計算：");
            sb.AppendLine(">");
            sb.AppendLine("> 個人追蹤數量 + 個人專案Star數量 + 個人專案Fork數量 + 組織貢獻專案的Star + 組織貢獻專案的Fork");
            sb.AppendLine(">");
            sb.AppendLine("> - 追蹤數 > 100");
            sb.AppendLine("> - 組織專案、其他專案貢獻分數公式 = 排名百分比 * star/fork = ((總人數-排名+1)/總人數) * star/fork");
            sb.AppendLine("> - 不能只有追蹤，沒有其他專案 star 或是 fork (要求其他分數加起來>10)");
            sb.AppendLine("> - 因為欄位有限，顯示只取前幾名專案，完整專案資料可以看 User.json 資料集");
            sb.AppendLine();
            sb.AppendLine($"**更新時間**: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
            sb.AppendLine($"**總計用戶數**: {users.Count}");
            sb.AppendLine();

            // 生成表格標題
            sb.AppendLine("| 排名 | Total Influence | 開發者 | Followers | Personal Projects | Top Org Projects | Top Contributed Projects |");
            sb.AppendLine("|------|-----------------|--------|-----------|-------------------|------------------|--------------------------|");

            for (int i = 0; i < users.Count; i++)
            {
                var user = users[i];
                var rank = i + 1;
                var totalInfluence = $"**{user.Score:F0}**";
                
                // 開發者資訊 (頭像 + 姓名 + 位置)
                var developerInfo = $"[<img src=\"{user.AvatarUrl}&s=32\" width=\"32\" height=\"32\" style=\"border-radius: 50%;\" />]({user.HtmlUrl})<br/>**[{user.Login}]({user.HtmlUrl})**<br/>{user.Name}";
                if (!string.IsNullOrEmpty(user.Location))
                {
                    developerInfo += $"<br/>📍 {user.Location}";
                }
                
                var followers = user.Followers.ToString("N0");
                
                // 個人專案資訊
                var personalProjects = "";
                if (user.TopRepositories.Any())
                {
                    var totalStars = user.TopRepositories.Sum(r => r.StargazersCount);
                    var totalForks = user.TopRepositories.Sum(r => r.ForksCount);
                    personalProjects = $"⭐ {totalStars:N0} 🍴 {totalForks:N0}<br/><br/>";
                    
                    var topRepos = user.TopRepositories.ToList();
                    for (int j = 0; j < topRepos.Count; j++)
                    {
                        var repo = topRepos[j];
                        personalProjects += $"• [{repo.Name}]({repo.HtmlUrl}) ({repo.StargazersCount:N0}⭐)";
                        if (j < topRepos.Count - 1)
                        {
                            personalProjects += "<br/>";
                        }
                    }
                }
                else
                {
                    personalProjects = "-";
                }
                
                // 組織貢獻專案資訊
                var orgContributedProjects = "";
                if (user.TopOrganizationRepositories.Any())
                {
                    var totalOrgStars = user.TopOrganizationRepositories.Sum(r => r.StargazersCount);
                    var totalOrgForks = user.TopOrganizationRepositories.Sum(r => r.ForksCount);
                    orgContributedProjects = $"⭐ {totalOrgStars:N0} 🍴 {totalOrgForks:N0}<br/><br/>";
                    
                    var topOrgRepos = user.TopOrganizationRepositories.ToList();
                    for (int j = 0; j < topOrgRepos.Count; j++)
                    {
                        var repo = topOrgRepos[j];
                        var repoDisplay = $"• [{repo.Name}]({repo.HtmlUrl}) ({repo.StargazersCount:N0}⭐)";
                        if (!string.IsNullOrEmpty(repo.RankDisplay))
                        {
                            repoDisplay += $" {repo.RankDisplay}";
                        }
                        orgContributedProjects += repoDisplay;
                        if (j < topOrgRepos.Count - 1)
                        {
                            orgContributedProjects += "<br/>";
                        }
                    }
                }
                else
                {
                    orgContributedProjects = "-";
                }
                
                // 其他個人專案貢獻資訊
                var otherContributedProjects = "";
                if (user.TopContributedRepositories.Any())
                {
                    var totalContribStars = user.TopContributedRepositories.Sum(r => r.StargazersCount);
                    var totalContribForks = user.TopContributedRepositories.Sum(r => r.ForksCount);
                    otherContributedProjects = $"⭐ {totalContribStars:N0} 🍴 {totalContribForks:N0}<br/>";
                    
                    var topContribRepos = user.TopContributedRepositories.ToList();
                    for (int j = 0; j < topContribRepos.Count; j++)
                    {
                        var repo = topContribRepos[j];
                        var repoDisplay = $"• [{repo.Name}]({repo.HtmlUrl}) ({repo.StargazersCount:N0}⭐)";
                        if (!string.IsNullOrEmpty(repo.RankDisplay))
                        {
                            repoDisplay += $" {repo.RankDisplay}";
                        }
                        otherContributedProjects += repoDisplay;
                        if (j < topContribRepos.Count - 1)
                        {
                            otherContributedProjects += "<br/>";
                        }
                    }
                }
                else
                {
                    otherContributedProjects = "-";
                }
                
                // 轉義管道符號以避免表格格式錯誤
                developerInfo = developerInfo.Replace("|", "\\|");
                personalProjects = personalProjects.Replace("|", "\\|");
                orgContributedProjects = orgContributedProjects.Replace("|", "\\|");
                otherContributedProjects = otherContributedProjects.Replace("|", "\\|");
                
                sb.AppendLine($"| {rank} | {totalInfluence} | {developerInfo} | {followers} | {personalProjects} | {orgContributedProjects} | {otherContributedProjects} |");
            }
            
            return sb.ToString();
        }

        static string GenerateHtml(List<GitHubUser> users)
        {
            var sb = new StringBuilder();
            sb.AppendLine("<!DOCTYPE html>");
            sb.AppendLine("<html lang=\"zh-Hant\">");
            sb.AppendLine("<head>");
            sb.AppendLine("    <meta charset=\"UTF-8\">");
            sb.AppendLine("    <meta name=\"viewport\" content=\"width=device-width, initial-scale=1.0\">");
            sb.AppendLine($"    <title>{currentRegion.ChineseName}知名GitHub用戶排名</title>");
            sb.AppendLine("    <link rel=\"stylesheet\" href=\"https://cdnjs.cloudflare.com/ajax/libs/normalize/8.0.1/normalize.min.css\">");
            sb.AppendLine("    <style>");
            sb.AppendLine("        body { font-family: 'Segoe UI', 'Noto Sans TC', Arial, sans-serif; background: #f7f7f7; color: #222; }");
            sb.AppendLine("        h1 { text-align: center; margin-top: 2rem; }");
            sb.AppendLine("        table { border-collapse: collapse; margin: 2rem auto; background: #fff; box-shadow: 0 2px 8px #0001; }");
            sb.AppendLine("        th, td { padding: 0.7rem 1rem; border: 1px solid #ddd; text-align: center; }");
            sb.AppendLine("        th { background: #222; color: #fff; }");
            sb.AppendLine("        tr:nth-child(even) { background: #f2f2f2; }");
            sb.AppendLine("        .avatar { border-radius: 50%; width: 32px; height: 32px; vertical-align: middle; }");
            sb.AppendLine("        .badge-btn { background: none; border: none; cursor: pointer; padding: 0; }");
            sb.AppendLine("        .rank-info { color: #666; font-size: 0.9em; font-style: italic; }");
            sb.AppendLine("        .project-stats { color: #555; font-weight: bold; }");
            sb.AppendLine("    </style>");
            sb.AppendLine("</head>");
            sb.AppendLine("<body>");
            sb.AppendLine($"<h1>{currentRegion.ChineseName}知名GitHub用戶排名</h1>");
            sb.AppendLine("<div style='text-align:center; margin-bottom: 2rem;'>");
            sb.AppendLine("    <a href='index.html' style='color: #3498db; text-decoration: none; margin: 0 15px; padding: 8px 16px; border: 1px solid #3498db; border-radius: 4px; background: #3498db; color: white;'>🏆 開發者排名</a>");
            sb.AppendLine("    <a href='taiwan-projects.html' style='color: #3498db; text-decoration: none; margin: 0 15px; padding: 8px 16px; border: 1px solid #3498db; border-radius: 4px;'>📂 專案排名</a>");
            sb.AppendLine("</div>");
            sb.AppendLine($"<p style='text-align:center;'>更新時間: {DateTime.Now:yyyy-MM-dd HH:mm:ss}｜總計用戶數: {users.Count}</p>");
            sb.AppendLine("<table>");
            sb.AppendLine("<tr><th>Score</th><th>排名</th><th>開發者</th><th>Followers</th><th>Personal Projects</th><th>Top Org Projects</th><th>Top Contributed Projects</th></tr>");
            for (int i = 0; i < users.Count; i++)
            {
                var user = users[i];
                var rank = i + 1;
                var badgeUrl = $"https://img.shields.io/badge/NO{rank}%20{user.Score:F0}_-red?style=for-the-badge&logo=github&logoColor=white&labelColor=black";
                var badgeHtml = $"<button class='badge-btn' onclick=\"navigator.clipboard.writeText('{badgeUrl}')\"><img src='{badgeUrl}' alt='K.O.榜戰力指數' title='點擊複製 badge 連結' /></button>";
                var developerInfo = $"<a href='{user.HtmlUrl}' target='_blank'><img class='avatar' src='{user.AvatarUrl}&s=32' alt='{user.Login}' /></a><br/><a href='{user.HtmlUrl}' target='_blank'><b>{user.Login}</b></a><br/>{user.Name}";
                if (!string.IsNullOrEmpty(user.Location))
                    developerInfo += $"<br/>📍 {user.Location}";
                var followers = user.Followers.ToString("N0");
                var personalProjects = "-";
                if (user.TopRepositories.Any())
                {
                    var totalStars = user.TopRepositories.Sum(r => r.StargazersCount);
                    var totalForks = user.TopRepositories.Sum(r => r.ForksCount);
                    personalProjects = $"<span class='project-stats'>⭐ {totalStars:N0} 🍴 {totalForks:N0}</span><br/>";
                    var topRepos = user.TopRepositories.ToList();
                    for (int j = 0; j < topRepos.Count; j++)
                    {
                        var repo = topRepos[j];
                        personalProjects += $"• <a href='{repo.HtmlUrl}' target='_blank'>{repo.Name}</a> ({repo.StargazersCount:N0}⭐)";
                        if (j < topRepos.Count - 1) personalProjects += "<br/>";
                    }
                }
                var contributedProjects = "-";
                if (user.TopOrganizationRepositories.Any())
                {
                    var totalOrgStars = user.TopOrganizationRepositories.Sum(r => r.StargazersCount);
                    var totalOrgForks = user.TopOrganizationRepositories.Sum(r => r.ForksCount);
                    contributedProjects = $"<span class='project-stats'>⭐ {totalOrgStars:N0} 🍴 {totalOrgForks:N0}</span><br/>";
                    var topOrgRepos = user.TopOrganizationRepositories.ToList();
                    for (int j = 0; j < topOrgRepos.Count; j++)
                    {
                        var repo = topOrgRepos[j];
                        var repoDisplay = $"• <a href='{repo.HtmlUrl}' target='_blank'>{repo.Name}</a> ({repo.StargazersCount:N0}⭐)";
                        
                        // 添加排名資訊
                        if (!string.IsNullOrEmpty(repo.RankDisplay))
                        {
                            repoDisplay += $"<br/>&nbsp;&nbsp;<span class='rank-info'>{repo.RankDisplay}</span>";
                        }
                        
                        contributedProjects += repoDisplay;
                        if (j < topOrgRepos.Count - 1) contributedProjects += "<br/>";
                    }
                }
                var otherContributedProjects = "-";
                if (user.TopContributedRepositories.Any())
                {
                    var totalContribStars = user.TopContributedRepositories.Sum(r => r.StargazersCount);
                    var totalContribForks = user.TopContributedRepositories.Sum(r => r.ForksCount);
                    otherContributedProjects = $"<span class='project-stats'>⭐ {totalContribStars:N0} 🍴 {totalContribForks:N0}</span><br/>";
                    var topContribRepos = user.TopContributedRepositories.ToList();
                    for (int j = 0; j < topContribRepos.Count; j++)
                    {
                        var repo = topContribRepos[j];
                        var repoDisplay = $"• <a href='{repo.HtmlUrl}' target='_blank'>{repo.Name}</a> ({repo.StargazersCount:N0}⭐)";
                        
                        // 添加排名資訊
                        if (!string.IsNullOrEmpty(repo.RankDisplay))
                        {
                            repoDisplay += $"<br/>&nbsp;&nbsp;<span class='rank-info'>{repo.RankDisplay}</span>";
                        }
                        
                        otherContributedProjects += repoDisplay;
                        if (j < topContribRepos.Count - 1) otherContributedProjects += "<br/>";
                    }
                }
                sb.AppendLine($"<tr><td>{badgeHtml}</td><td>{rank}</td><td>{developerInfo}</td><td>{followers}</td><td>{personalProjects}</td><td>{contributedProjects}</td><td>{otherContributedProjects}</td></tr>");
            }
            sb.AppendLine("</table>");
            sb.AppendLine("<p style='text-align:center;color:#888;'>點擊 badge 可複製 badge 連結，可用於個人 README 或其他地方展示。</p>");
            sb.AppendLine("</body></html>");
            return sb.ToString();
        }
    }
}
