namespace IfcToExcelWinForms
{
    public record FaceRow(
        string Face,
        string Active,
        string? CutDistanceFromColumn,
        string? CutProfile,
        string? BoltEdgeDistance,
        string? BoltTopDistance,
        string? BoltSpacing,
        string? SwapSide,
        string? CutRoundingRadius,
        string? AddSupportAngle,
        string? AngleLength,
        string? AngleProfile,
        string? AngleBoltEdgeDistance,
        string? AngleBoltTopDistance,
        string? AngleBoltSpacing,
        string? BoltSize,
        string? AnglePartNumber,
        string? AngleRoomLetter,
        string? AngleOffsetFromBeamEnd,
        string? OffsetFromEdgeOfColumn,
        string? BeamEndOffsetDistance
    );
}