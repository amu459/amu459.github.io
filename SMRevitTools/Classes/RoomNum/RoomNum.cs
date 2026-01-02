using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using SMRevitTools.Classes.MeetCeiling;

namespace SMRevitTools
{
    [TransactionAttribute(TransactionMode.Manual)]
    public class RoomNum : IExternalCommand
    {
        // Maps full room name to standard short code
        static readonly Dictionary<string, string> roomDict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            {"ARRIVAL AREA", "ARVL"},
            {"CAFETERIA", "CAFE"},
            {"REPRO", "RPR"},
            {"PANTRY", "PAN"},
            {"PHONEBOOTH", "PHB"},
            {"PHONEROOM", "PHR"},
            {"1 PAX FOCUS", "1PF"},
            {"1 PAX MEETING", "1PAX"},
            {"1 PAX PHONEROOM", "1PHR"},
            {"1 PAX PHONEBOOTH", "1PHB"},
            {"2 PAX FOCUS", "2PF"},
            {"2 PAX PHONE ROOM", "2PHR"},
            {"2 PAX PHONEBOOTH", "2PHB"},
            {"2 PAX MEETING", "2PAX"},
            {"3 PAX MEETING", "3PAX"},
            {"4 PAX MEETING", "4PAX"},
            {"5 PAX MEETING", "5PAX"},
            {"6 PAX MEETING", "6PAX"},
            {"7 PAX MEETING", "7PAX"},
            {"8 PAX MEETING", "8PAX"},
            {"9 PAX MEETING", "9PAX"},
            {"10 PAX MEETING", "10PAX"},
            {"11 PAX MEETING", "11PAX"},
            {"12 PAX MEETING", "12PAX"},
            {"13 PAX MEETING", "13PAX"},
            {"14 PAX MEETING", "14PAX"},
            {"15 PAX MEETING", "15PAX"},
            {"16 PAX MEETING", "16PAX"},
            {"17 PAX MEETING", "17PAX"},
            {"18 PAX MEETING", "18PAX"},
            {"BOARDROOM", "BOARD"},
            {"CONFERENCE ROOM", "CONF" },
            {"TRAINING ROOM", "TRN"},
            {"LOUNGE MEETING", "LMT"},
            {"CABIN", "CBN"},
            {"EXECUTIVE CABIN", "EXECB"},
            {"MAIL ROOM", "MAIL"},
            {"CLOAK ROOM", "CLKRM"},
            {"STORE ROOM", "STR"},
            {"GAME ROOM", "GAME"},
            {"ELECTRICAL ROOM", "ELEC"},
            {"MECHANICAL ROOM", "MECH"},
            {"IDF ROOM", "IDF"},
            {"MDF ROOM", "MDF"},
            {"IT ROOM", "IT"},
            {"AHU", "AHU"},
            {"SECURITY ROOM", "SCRT"},
            {"BMS", "BMS"},
            {"PRAYER ROOM", "PRRM"},
            {"COFFEE POINT", "COFP"},
            {"MEDICAL ROOM", "MEDC"},
            {"SNOOZE ROOM", "SNZ"},
            {"TEA POINT", "TEA"},
            {"CHAI POINT", "CHAI"},
            {"SERVER ROOM", "SER"},
            {"UPS BATTERY ROOM", "UPS"},
            {"TOILET", "TOLT"},
            {"MALE TOILET", "MT"},
            {"FEMALE TOILET", "FT"},
            {"MEN TOILET", "MT"},
            {"WOMEN TOILET", "WT"},
            {"PD TOILET", "PDT"},
            {"ACCESSIBLE TOILET", "ACCT" },
            {"WASHROOM", "WSHR"},
            {"SHOWER", "SHWR"},
            {"PD WASHROOM", "PDWR"},
            {"WELLNESS ROOM", "WELL"},
            {"MOTHERS ROOM", "WELL"},
            {"FACILITY ROOM", "FCLT"},
            {"MOP ROOM", "MOP"},
            {"JANITOR", "JAN"},
            {"CHANGING ROOM", "CHNG"},
            {"HUB ROOM", "HUB"},
            {"BATTERY", "BAT"},
            {"UPS", "UPS"},
            {"CONVERSATION ROOM", "CONV"},
            {"CEO CABIN", "CEO"},
            {"LOUNGE", "LO"},
            {"COLLAB", "CLB"},
            {"COLLAB AREA", "CLBA"},
            {"FOCUS", "FO"},
            {"WORK", "WRK"},
            {"WORK STATION", "WRK"},
            {"RECEPTION", "RECP"},
            {"WAREHOUSE", "WARE" },
            {"CORRIDOR", "COR" },
            {"PASSAGE", "PASS" },
            {"OFFICE", "OFF"}
        };

        // Tracks assigned suffixes per (Level + RoomCode)
        static Dictionary<string, HashSet<int>> roomCounters;

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            DateTime startTime = DateTime.Now;
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                // Get selected rooms
                List<Room> rooms = RoomMethods.GetRoomsFromSelection(uidoc);
                if (rooms.Count == 0)
                {
                    message = "No rooms selected.";
                    return Result.Failed;
                }
                rooms = rooms.Where(r =>
                {
                    string fullName = r.Name?.Trim() ?? "";
                    string number = r.Number?.Trim() ?? "";

                    // Remove room number from name if it ends with it
                    if (!string.IsNullOrEmpty(number) && fullName.EndsWith(number))
                    {
                        fullName = fullName.Substring(0, fullName.Length - number.Length).Trim();
                    }

                    // Now fullName should contain only the room's actual name portion
                    return !string.Equals(fullName, "Room", StringComparison.OrdinalIgnoreCase);
                }).ToList();

                // Get selected room Ids to exclude them from existing suffix count
                IList<ElementId> selectedRoomIds = rooms.Select(r => r.Id).ToList();

                // Initialize counters excluding selected rooms
                roomCounters = InitializeRoomCounters(doc, selectedRoomIds);

                // Get room names for selected rooms and find closest room type code using Levenshtein distance
                List<string> roomNames = RoomMethods.GetRoomNames(rooms);
                List<string> roomCodes = new List<string>();

                foreach (string roomName in roomNames)
                {
                    string bestMatch = null;
                    int bestScore = int.MaxValue;

                    foreach (string room in roomDict.Keys)
                    {
                        int dist = LevenshteinDistance(roomName.ToUpper().Trim(), room.ToUpper());
                        if (dist < bestScore)
                        {
                            bestScore = dist;
                            bestMatch = room;
                        }
                    }

                    if (bestMatch != null && bestScore <20)
                    {
                        roomCodes.Add(roomDict[bestMatch]);
                    }
                    else
                    {
                        roomCodes.Add(roomName.Substring(0,2));
                    }
                }

                using (Transaction transaction = new Transaction(doc, "Assign Room Numbers"))
                {
                    transaction.Start();

                    int count = 0;
                    foreach (Room room in rooms)
                    {
                        // Normalize level code
                        string geometricLevel = room.Level.LookupParameter("SM-Geometric Level")?.AsString() ?? "UNKNOWN";

                        if (!string.IsNullOrEmpty(geometricLevel))
                        {
                            if (char.IsDigit(geometricLevel[0]))
                            {
                                // Pad to two digits: "1" → "F01", "10" → "F10"
                                if (geometricLevel.Length == 1)
                                    geometricLevel = "F0" + geometricLevel;
                                else
                                    geometricLevel = "F" + geometricLevel;
                            }
                            else
                            {
                                geometricLevel = geometricLevel.ToUpper();
                            }
                        }

                        // Generate a new suffix number ignoring existing suffix on this room
                        string newRoomNumber = GenerateRoomNumber(geometricLevel, roomCodes[count], roomCounters);

                        room.Number = newRoomNumber;
                        count++;
                    }

                    transaction.Commit();
                }

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }

        // Initialize roomCounters excluding selected rooms
        static Dictionary<string, HashSet<int>> InitializeRoomCounters(Document doc, IList<ElementId> excludedRoomIds)
        {
            var counters = new Dictionary<string, HashSet<int>>();

            var collector = new FilteredElementCollector(doc)
                .OfClass(typeof(SpatialElement))
                .OfType<Room>();

            foreach (Room existingRoom in collector)
            {
                // Skip rooms that are in selectedRoomIds (will be renumbered)
                if (excludedRoomIds.Contains(existingRoom.Id))
                    continue;

                var parsed = ParseRoomNumber(existingRoom.Number);
                if (parsed != null)
                {
                    string key = parsed.Value.levelCode.ToUpper() + "-" + parsed.Value.roomCode.ToUpper();
                    int suffix = parsed.Value.suffix;

                    if (!counters.ContainsKey(key))
                        counters[key] = new HashSet<int>();

                    counters[key].Add(suffix);
                }
            }

            return counters;
        }

        // Parse room number format "F18-4PAX-03" into components
        static (string levelCode, string roomCode, int suffix)? ParseRoomNumber(string roomNumber)
        {
            if (string.IsNullOrEmpty(roomNumber))
                return null;

            string[] parts = roomNumber.Split('-');
            if (parts.Length < 3)
                return null;

            if (!int.TryParse(parts[2], out int suffix))
                return null;

            return (parts[0], parts[1], suffix);
        }

        // Generate smallest available suffix (starting 1) ignoring suffixes taken by unselected rooms
        static string GenerateRoomNumber(string levelCode, string roomShortCode, Dictionary<string, HashSet<int>> counters)
        {
            string key = levelCode.ToUpper() + "-" + roomShortCode.ToUpper();

            if (!counters.ContainsKey(key))
                counters[key] = new HashSet<int>();

            int suffix = 1;
            while (counters[key].Contains(suffix))
            {
                suffix++;
            }

            counters[key].Add(suffix);

            return $"{levelCode}-{roomShortCode}-{suffix.ToString("D2")}";
        }

        // Levenshtein distance for fuzzy matching
        public static int LevenshteinDistance(string s, string t)
        {
            int n = s.Length;
            int m = t.Length;
            int[,] d = new int[n + 1, m + 1];

            if (n == 0) return m;
            if (m == 0) return n;

            for (int i = 0; i <= n; d[i, 0] = i++) { }
            for (int j = 0; j <= m; d[0, j] = j++) { }

            for (int i = 1; i <= n; i++)
                for (int j = 1; j <= m; j++)
                {
                    int cost = (t[j - 1] == s[i - 1]) ? 0 : 1;
                    d[i, j] = Math.Min(
                        Math.Min(d[i - 1, j] + 1, d[i, j - 1] + 1),
                        d[i - 1, j - 1] + cost);
                }
            return d[n, m];
        }
    }
}
