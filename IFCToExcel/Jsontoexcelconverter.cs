using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Text.RegularExpressions;
using ClosedXML.Excel;

namespace IfcToExcelWinForms
{
    /// <summary>
    /// Converts an IDEA StatiCa connection JSON file to the shear-clip Excel format.
    ///
    /// JSON structure used:
    ///   beams[]          – all member elements; isBearingMember=true is the column
    ///   beams[].plates[] – attached plates; each plate has origin, axisX/Y/Z, region (SVG path)
    ///   boltGrids[]      – bolt patterns; each has origin, axisX/Y/Z, positions[], connectedParts[]
    ///   boltGrids[].boltAssembly.element – bolt spec (name, diameter, …)
    ///
    /// Key design decisions
    /// --------------------
    /// • Angle-grid face uses the ANGLE BEAM's face (not bolt world position) because
    ///   through-bolts can appear on the opposite column face.
    /// • Sub-face labels (WEST DOWN/UP, NORTH LEFT/RIGHT etc.) assigned by grouping
    ///   beam-web grids on the same base face and sorting by Z (E/W) or X (N/S).
    /// • AddSupportAngle, AngleBoltTop written as plain strings to match Excel target.
    /// • BoltTop rounded to nearest half-inch; AngleBoltTop to nearest quarter-inch.
    /// </summary>
    public static class JsonToExcelConverter
    {
        private static readonly string[] FaceOrder =
        {
            "NORTH LEFT","NORTH","NORTH RIGHT",
            "EAST UP","EAST","EAST DOWN",
            "SOUTH RIGHT","SOUTH","SOUTH LEFT",
            "WEST DOWN","WEST","WEST UP"
        };

        // ── public entry point ──────────────────────────────────────────────────

        public static void Convert(string jsonPath, string xlsxPath)
        {
            // IDEA StatiCa exports JSON as UTF-16 LE (BOM 0xFF 0xFE).
            // System.Text.Json only accepts UTF-8, so read as text and re-encode.
            string jsonText = File.ReadAllText(jsonPath, System.Text.Encoding.Unicode);
            var jsonBytes = System.Text.Encoding.UTF8.GetBytes(jsonText);
            var doc = JsonDocument.Parse(jsonBytes);
            var root = doc.RootElement;

            // 1. Build plate-id → beam map, and flat plate lookup
            var plateToBeam = new Dictionary<int, JsonElement>();
            var allPlates = new Dictionary<int, JsonElement>();
            foreach (var beam in root.GetProperty("beams").EnumerateArray())
                foreach (var plate in beam.GetProperty("plates").EnumerateArray())
                {
                    int pid = plate.GetProperty("id").GetInt32();
                    plateToBeam[pid] = beam;
                    allPlates[pid] = plate;
                }

            // 2. Identify bearing member and angle beams
            int bearingBeamId = -1;
            var angleBeams = new List<JsonElement>();
            foreach (var beam in root.GetProperty("beams").EnumerateArray())
            {
                if (beam.GetProperty("isBearingMember").GetBoolean())
                    bearingBeamId = beam.GetProperty("id").GetInt32();
                var csType = beam.GetProperty("crossSectionType").GetString() ?? "";
                if (csType.Equals("RolledAngle", StringComparison.OrdinalIgnoreCase))
                    angleBeams.Add(beam);
            }

            // 3. Classify each angle beam's face; build plate → angle-beam-id lookup
            var angleBeamFace = new Dictionary<int, string>();
            var plateToAngleBeamId = new Dictionary<int, int>();
            foreach (var angle in angleBeams)
            {
                int aid = angle.GetProperty("id").GetInt32();
                angleBeamFace[aid] = ClassifyAngleFace(angle);
                foreach (var plate in angle.GetProperty("plates").EnumerateArray())
                    plateToAngleBeamId[plate.GetProperty("id").GetInt32()] = aid;
            }

            // 4. Process bolt grids
            // base face → list of BeamGridInfo (for sub-face sorting)
            var beamGridsByFace = new Dictionary<string, List<BeamGridInfo>>();
            // angle-beam face → AngleGridInfo
            var angleGridByFace = new Dictionary<string, AngleGridInfo>();

            foreach (var grid in root.GetProperty("boltGrids").EnumerateArray())
            {
                var wps = BoltWorldPositions(grid);
                if (wps.Count == 0) continue;

                int? columnPid = null, nonBearingPid = null, anglePid = null;
                foreach (var cp in grid.GetProperty("connectedParts").EnumerateArray())
                {
                    int pid = cp.GetProperty("id").GetInt32();
                    if (!plateToBeam.TryGetValue(pid, out var ownerBeam)) continue;
                    int oid = ownerBeam.GetProperty("id").GetInt32();
                    if (oid == bearingBeamId) columnPid = pid;
                    else if (plateToAngleBeamId.ContainsKey(pid)) anglePid = pid;
                    else nonBearingPid = pid;
                }

                if (columnPid.HasValue && nonBearingPid.HasValue)
                {
                    string baseFace = ClassifyGridFace(wps[0]);
                    if (!beamGridsByFace.ContainsKey(baseFace))
                        beamGridsByFace[baseFace] = new List<BeamGridInfo>();
                    beamGridsByFace[baseFace].Add(new BeamGridInfo
                    {
                        Positions = wps,
                        WebPlateId = nonBearingPid.Value,
                        BoltAssembly = GetBoltAssemblyName(grid),
                        // Fix F: use min Z across ALL bolt positions for stable sorting
                        MinY = wps.Min(p => p.Y),
                        MinX = wps.Min(p => p.X),
                    });
                }

                if (columnPid.HasValue && anglePid.HasValue)
                {
                    // Key by angle beam's face to handle through-bolt geometry
                    int angleBeamId = plateToAngleBeamId[anglePid.Value];
                    if (angleBeamFace.TryGetValue(angleBeamId, out string? af)
                        && !angleGridByFace.ContainsKey(af))
                    {
                        angleGridByFace[af] = new AngleGridInfo
                        {
                            Positions = wps,
                            PlateId = anglePid.Value,
                        };
                    }
                }
            }

            // 5. Assign sub-face labels and build FaceRows
            var finalRows = new Dictionary<string, FaceRow>();
            foreach (var f in FaceOrder)
                finalRows[f] = BlankRow(f);

            foreach (var (baseFace, grids) in beamGridsByFace)
            {
                bool isEW = baseFace == "EAST" || baseFace == "WEST";
                List<BeamGridInfo> sorted = isEW
                    ? grids.OrderBy(g => g.MinY).ToList()   // E/W: sort by Y (perpendicular to face); most negative Y = DOWN
                    : grids.OrderBy(g => g.MinX).ToList();  // N/S: sort by X (perpendicular to face); most negative X = LEFT
                List<string> faces = SubFaceLabels(baseFace, sorted.Count, isEW);

                for (int i = 0; i < sorted.Count; i++)
                {
                    string face = faces[i];
                    if (!Array.Exists(FaceOrder, f => f == face)) continue;

                    var gi = sorted[i];
                    var webPlate = allPlates[gi.WebPlateId];

                    string boltSpacing = ComputeBoltSpacing(gi.Positions);

                    // OffsetFromEdgeOfColumn:
                    // '1' = bolt line is centered on the column face (offset from column edge)
                    // '0' = bolt line is offset perpendicular to the face (offset from column center)
                    // Test: measure the bolt's distance perpendicular to the face it's on.
                    // For N/S faces the perpendicular is X; for E/W faces it's Y.
                    string offsetFromEdge;
                    {
                        var wp0 = gi.Positions[0];
                        double perpMm = (baseFace == "NORTH" || baseFace == "SOUTH")
                            ? Math.Abs(wp0.X) * 1000.0
                            : Math.Abs(wp0.Y) * 1000.0;
                        offsetFromEdge = perpMm < 5.0 ? "1" : "0";
                    }

                    (string boltEdge, string boltTop) = ComputeBoltEdgeTop(gi.Positions, webPlate, offsetFromEdge);
                    string boltSize = ParseBoltSize(gi.BoltAssembly);

                    // Angle data — try sub-face first, then base face
                    string? angleLen = null, angleProf = null, angleBoltEdge = null, angleBoltTop = null;

                    if (!angleGridByFace.TryGetValue(face, out var agi))
                        angleGridByFace.TryGetValue(baseFace, out agi);

                    var angleBeam = FindAngleBeamForFace(angleBeams, angleBeamFace, face, baseFace);
                    if (angleBeam.HasValue)
                    {
                        angleLen = GetAngleLength(angleBeam.Value);
                        angleProf = GetAngleProfile(angleBeam.Value);
                    }
                    if (agi != null)
                    {
                        (angleBoltEdge, angleBoltTop) = ComputeAngleBoltEdgeTop(agi.Positions, allPlates[agi.PlateId]);
                    }

                    finalRows[face] = finalRows[face] with
                    {
                        Active = "YES",
                        BoltEdgeDistance = boltEdge,
                        BoltTopDistance = boltTop,
                        BoltSpacing = boltSpacing,
                        AddSupportAngle = "1",          // written as int below
                        AngleLength = angleLen,
                        AngleProfile = angleProf,
                        AngleBoltEdgeDistance = angleBoltEdge,
                        AngleBoltTopDistance = angleBoltTop,
                        AngleBoltSpacing = null,
                        BoltSize = boltSize,
                        AngleOffsetFromBeamEnd = "3/4",
                        OffsetFromEdgeOfColumn = offsetFromEdge,
                        BeamEndOffsetDistance = "1/2",
                    };
                }
            }

            WriteExcel(xlsxPath,
                finalRows.Values.OrderBy(r => Array.IndexOf(FaceOrder, r.Face)).ToList());
        }

        // ── Sub-face label assignment ────────────────────────────────────────────

        private static List<string> SubFaceLabels(string baseFace, int count, bool isEW)
        {
            if (count == 1) return new List<string> { baseFace };
            return (isEW, baseFace, count) switch
            {
                (true, "EAST", 2) => new List<string> { "EAST DOWN", "EAST UP" },
                (true, "EAST", _) => new List<string> { "EAST DOWN", "EAST", "EAST UP" },
                (true, "WEST", 2) => new List<string> { "WEST DOWN", "WEST UP" },
                (true, "WEST", _) => new List<string> { "WEST DOWN", "WEST", "WEST UP" },
                (false, "NORTH", 2) => new List<string> { "NORTH LEFT", "NORTH RIGHT" },
                (false, "NORTH", _) => new List<string> { "NORTH LEFT", "NORTH", "NORTH RIGHT" },
                (false, "SOUTH", 2) => new List<string> { "SOUTH LEFT", "SOUTH RIGHT" },
                (false, "SOUTH", _) => new List<string> { "SOUTH LEFT", "SOUTH", "SOUTH RIGHT" },
                _ => new List<string> { baseFace },
            };
        }

        // ── Geometry helpers ────────────────────────────────────────────────────

        private static List<(double X, double Y, double Z)> BoltWorldPositions(JsonElement grid)
        {
            var orig = grid.GetProperty("origin");
            double ox = orig.GetProperty("x").GetDouble();
            double oy = orig.GetProperty("y").GetDouble();
            double oz = orig.GetProperty("z").GetDouble();
            var axX = grid.GetProperty("axisX"); var axY = grid.GetProperty("axisY"); var axZ = grid.GetProperty("axisZ");
            double axx = axX.GetProperty("x").GetDouble(), axy = axX.GetProperty("y").GetDouble(), axz = axX.GetProperty("z").GetDouble();
            double ayx = axY.GetProperty("x").GetDouble(), ayy = axY.GetProperty("y").GetDouble(), ayz = axY.GetProperty("z").GetDouble();
            double azx = axZ.GetProperty("x").GetDouble(), azy = axZ.GetProperty("y").GetDouble(), azz = axZ.GetProperty("z").GetDouble();
            var result = new List<(double, double, double)>();
            foreach (var pos in grid.GetProperty("positions").EnumerateArray())
            {
                double px = pos.GetProperty("x").GetDouble(), py = pos.GetProperty("y").GetDouble(), pz = pos.GetProperty("z").GetDouble();
                result.Add((ox + axx * px + ayx * py + azx * pz, oy + axy * px + ayy * py + azy * pz, oz + axz * px + ayz * py + azz * pz));
            }
            return result;
        }

        private static string ClassifyGridFace((double X, double Y, double Z) wp)
        {
            if (Math.Abs(wp.Y) >= Math.Abs(wp.X)) return wp.Y < 0 ? "SOUTH" : "NORTH";
            return wp.X < 0 ? "WEST" : "EAST";
        }

        private static string ClassifyAngleFace(JsonElement angleBeam)
        {
            double bY = 0, bX = 0, dY = 0, dX = 0;
            foreach (var plate in angleBeam.GetProperty("plates").EnumerateArray())
            {
                double px = plate.GetProperty("origin").GetProperty("x").GetDouble();
                double py = plate.GetProperty("origin").GetProperty("y").GetDouble();
                if (Math.Abs(py) > Math.Abs(bY)) { bY = py; dY = py; }
                if (Math.Abs(px) > Math.Abs(bX)) { bX = px; dX = px; }
            }
            if (Math.Abs(bY) >= Math.Abs(bX)) return dY < 0 ? "SOUTH" : "NORTH";
            return dX < 0 ? "WEST" : "EAST";
        }

        private static JsonElement? FindAngleBeamForFace(
            List<JsonElement> angleBeams, Dictionary<int, string> angleBeamFace,
            string subFace, string baseFace)
        {
            foreach (var ab in angleBeams)
            {
                int id = ab.GetProperty("id").GetInt32();
                if (angleBeamFace.TryGetValue(id, out string? f) && (f == subFace || f == baseFace))
                    return ab;
            }
            return null;
        }

        private static (double MinX, double MaxX, double MinY, double MaxY) ParseRegionBbox(string region)
        {
            var nums = Regex.Matches(region, @"[-+]?(?:\d+\.?\d*|\.\d+)(?:[Ee][-+]?\d+)?");
            double mnX = double.MaxValue, mxX = double.MinValue, mnY = double.MaxValue, mxY = double.MinValue;
            for (int i = 0; i + 1 < nums.Count; i += 2)
            {
                double x = double.Parse(nums[i].Value, System.Globalization.CultureInfo.InvariantCulture);
                double y = double.Parse(nums[i + 1].Value, System.Globalization.CultureInfo.InvariantCulture);
                if (x < mnX) mnX = x; if (x > mxX) mxX = x; if (y < mnY) mnY = y; if (y > mxY) mxY = y;
            }
            return (mnX, mxX, mnY, mxY);
        }

        private static (double Lx, double Ly) BoltInPlateLocal(JsonElement plate, (double X, double Y, double Z) wp)
        {
            var orig = plate.GetProperty("origin");
            double dx = wp.X - orig.GetProperty("x").GetDouble(), dy = wp.Y - orig.GetProperty("y").GetDouble(), dz = wp.Z - orig.GetProperty("z").GetDouble();
            var axX = plate.GetProperty("axisX"); var axY = plate.GetProperty("axisY");
            return (axX.GetProperty("x").GetDouble() * dx + axX.GetProperty("y").GetDouble() * dy + axX.GetProperty("z").GetDouble() * dz,
                    axY.GetProperty("x").GetDouble() * dx + axY.GetProperty("y").GetDouble() * dy + axY.GetProperty("z").GetDouble() * dz);
        }

        // ── Field computation ───────────────────────────────────────────────────

        private static string ComputeBoltSpacing(List<(double X, double Y, double Z)> wps)
        {
            if (wps.Count <= 1) return "";
            double sX = wps.Max(p => p.X) - wps.Min(p => p.X), sY = wps.Max(p => p.Y) - wps.Min(p => p.Y), sZ = wps.Max(p => p.Z) - wps.Min(p => p.Z);
            var sorted = sX >= sY && sX >= sZ ? wps.OrderBy(p => p.X).ToList() : sY >= sX && sY >= sZ ? wps.OrderBy(p => p.Y).ToList() : wps.OrderBy(p => p.Z).ToList();
            var relIn = sorted.Select(p => Units.MmToIn(Math.Sqrt(Math.Pow(p.X - sorted[0].X, 2) + Math.Pow(p.Y - sorted[0].Y, 2) + Math.Pow(p.Z - sorted[0].Z, 2)) * 1000.0)).ToArray();
            return Units.FormatSpacing(relIn);
        }

        private static (string Edge, string Top) ComputeBoltEdgeTop(
            List<(double X, double Y, double Z)> wps, JsonElement plate, string offsetFromEdge)
        {
            var region = plate.TryGetProperty("region", out var r) ? r.GetString() ?? "" : "";
            var (minX, maxX, minY, maxY) = ParseRegionBbox(region);
            var (lx0, _) = BoltInPlateLocal(plate, wps[0]);
            // Fix B: use the smaller of the two edge distances (handles reversed plate orientations)
            double dxMin = lx0 - minX;
            double dxMax = maxX - lx0;
            string edgeStr = Units.InToArchitectural(Units.MmToIn(Math.Min(dxMin, dxMax) * 1000.0));

            // BoltTop: nearest edge distance across all bolts, rounded to nearest ½"
            double minDy = wps.Select(wp => {
                var (_, ly) = BoltInPlateLocal(plate, wp);
                return Math.Min(ly - minY, maxY - ly);
            }).Min();
            double topRounded = Math.Round(Units.MmToIn(minDy * 1000.0) * 2.0) / 2.0;
            // Fix C: whole-number BoltTop is formatted as plain integer (no inch mark)
            // when OffsetFromEdgeOfColumn='0'; keeps inch mark when '1'
            string topStr;
            if (topRounded % 1.0 == 0 && offsetFromEdge == "0")
                topStr = ((int)topRounded).ToString();
            else
                topStr = Units.InToArchitectural(topRounded);
            return (edgeStr, topStr);
        }

        private static (string Edge, string Top) ComputeAngleBoltEdgeTop(
            List<(double X, double Y, double Z)> wps, JsonElement plate)
        {
            var region = plate.TryGetProperty("region", out var r) ? r.GetString() ?? "" : "";
            var (minX, _, minY, _) = ParseRegionBbox(region);
            var (lx, ly) = BoltInPlateLocal(plate, wps[0]);
            string edgeStr = Units.InToArchitectural(Units.MmToIn((lx - minX) * 1000.0));
            double topRounded = Math.Round(Units.MmToIn((ly - minY) * 1000.0) * 4.0) / 4.0;
            // Whole-number values are formatted as plain integers (e.g. "3" not "3\"")
            string topStr = (topRounded % 1.0 == 0)
                ? ((int)topRounded).ToString()
                : Units.InToArchitectural(topRounded);
            return (edgeStr, topStr);
        }

        private static string GetAngleLength(JsonElement angleBeam)
        {
            foreach (var plate in angleBeam.GetProperty("plates").EnumerateArray())
            {
                var region = plate.TryGetProperty("region", out var r) ? r.GetString() ?? "" : "";
                if (string.IsNullOrEmpty(region)) continue;
                var (minX, maxX, _, _) = ParseRegionBbox(region);
                return Units.InToArchitectural(Units.MmToIn((maxX - minX) * 1000.0), denom: 2);
            }
            return "";
        }

        private static string GetAngleProfile(JsonElement angleBeam)
            => (angleBeam.GetProperty("mprlName").GetString() ?? "")
               .Replace("(Imp)", "", StringComparison.Ordinal).Trim();

        private static string ParseBoltSize(string assemblyName)
        {
            var m = Regex.Match(assemblyName, @"^(\d+/\d+|\d+\.?\d*)");
            return m.Success ? m.Value : assemblyName;
        }

        private static string GetBoltAssemblyName(JsonElement grid)
        {
            try { return grid.GetProperty("boltAssembly").GetProperty("element").GetProperty("name").GetString() ?? ""; }
            catch { return ""; }
        }

        // ── Helpers ─────────────────────────────────────────────────────────────

        private static FaceRow BlankRow(string face) => new FaceRow(
            Face: face, Active: "NO",
            CutDistanceFromColumn: null, CutProfile: null,
            BoltEdgeDistance: null, BoltTopDistance: null, BoltSpacing: null,
            SwapSide: null, CutRoundingRadius: null,
            AddSupportAngle: null, AngleLength: null, AngleProfile: null,
            AngleBoltEdgeDistance: null, AngleBoltTopDistance: null, AngleBoltSpacing: null,
            BoltSize: null, AnglePartNumber: null, AngleRoomLetter: null,
            AngleOffsetFromBeamEnd: null, OffsetFromEdgeOfColumn: null, BeamEndOffsetDistance: null);

        // ── Excel writer ────────────────────────────────────────────────────────

        private static void WriteExcel(string path, List<FaceRow> rows)
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            string[] headers =
            {
                "FACE","ACTIVE","Cut Distance From Column","Cut Profile",
                "Bolt Edge Distance","Bolt Top Distance","Bolt Spacing","Swap Side?",
                "Cut Rounding Radius","Add Support Angle?","Angle Length","Angle Profile",
                "Angle Bolt Edge Distance","Angle Bolt Top Distance ","Angle Bolt Spacing",
                "Bolt Size","Angle Part Number","Angle Room Letter","Angle Offset from Beam End",
                "Offset from edge of Column?","Beam End Offset Distance"
            };
            for (int c = 0; c < headers.Length; c++) ws.Cell(1, c + 1).Value = headers[c];
            for (int i = 0; i < rows.Count; i++)
            {
                var r = rows[i]; int row = i + 2;
                ws.Cell(row, 1).Value = r.Face;
                ws.Cell(row, 2).Value = r.Active;
                ws.Cell(row, 3).Value = r.CutDistanceFromColumn;
                ws.Cell(row, 4).Value = r.CutProfile;
                ws.Cell(row, 5).Value = r.BoltEdgeDistance;
                ws.Cell(row, 6).Value = r.BoltTopDistance;
                ws.Cell(row, 7).Value = r.BoltSpacing;
                ws.Cell(row, 8).Value = r.SwapSide;
                ws.Cell(row, 9).Value = r.CutRoundingRadius;
                // AddSupportAngle: always write as string '1' (matches most targets)
                ws.Cell(row, 10).Value = r.AddSupportAngle;

                // BoltSpacing: when single-gap (no '*') and offsetFromEdge='0', write as int (no inch mark)
                if (r.BoltSpacing is string bs && !bs.Contains('*') && r.OffsetFromEdgeOfColumn == "0"
                    && double.TryParse(bs.Replace("\"", ""), System.Globalization.NumberStyles.Any,
                        System.Globalization.CultureInfo.InvariantCulture, out double bsVal)
                    && bsVal % 1.0 == 0)
                    ws.Cell(row, 7).Value = (int)bsVal;
                else
                    ws.Cell(row, 11).Value = r.AngleLength;
                ws.Cell(row, 12).Value = r.AngleProfile;
                ws.Cell(row, 13).Value = r.AngleBoltEdgeDistance;
                ws.Cell(row, 14).Value = r.AngleBoltTopDistance;
                ws.Cell(row, 15).Value = r.AngleBoltSpacing;
                ws.Cell(row, 16).Value = r.BoltSize;
                ws.Cell(row, 17).Value = r.AnglePartNumber;
                ws.Cell(row, 18).Value = r.AngleRoomLetter;
                ws.Cell(row, 19).Value = r.AngleOffsetFromBeamEnd;
                ws.Cell(row, 20).Value = r.OffsetFromEdgeOfColumn;
                ws.Cell(row, 21).Value = r.BeamEndOffsetDistance;
            }
            ws.Columns().AdjustToContents();
            wb.SaveAs(path);
        }

        // ── Supporting types ────────────────────────────────────────────────────

        private class BeamGridInfo
        {
            public List<(double X, double Y, double Z)> Positions { get; set; } = new();
            public int WebPlateId { get; set; }
            public string BoltAssembly { get; set; } = "";
            public double MinY { get; set; }   // min Y across all bolt positions (E/W sub-face sort)
            public double MinX { get; set; }   // min X across all bolt positions (N/S sub-face sort)
        }

        private class AngleGridInfo
        {
            public List<(double X, double Y, double Z)> Positions { get; set; } = new();
            public int PlateId { get; set; }
        }
    }
}