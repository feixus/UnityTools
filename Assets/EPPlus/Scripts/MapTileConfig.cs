using System;
using System.Collections.Generic;
using UnityEngine;

[CreateAssetMenu(fileName = "MapTileConfig", menuName = "TestScriptableObject", order = 1)]
public class MapTileConfig : ScriptableObject 
{
    public List<SurfaceTile> BasicTileData;
    public List<BuildingTile> BuildingData;

    public MapTileConfig()
    {
        BasicTileData = new List<SurfaceTile>();
        BuildingData = new List<BuildingTile>();
    }

}

// jsonUtility.fromjson 不支持ScriptableObject
public class MapTileConfigForJson
{
    public List<SurfaceTile> BasicTileData;
    public List<BuildingTile> BuildingData;

    public MapTileConfigForJson()
    {
        BasicTileData = new List<SurfaceTile>();
        BuildingData = new List<BuildingTile>();
    }

}

[Serializable]
public class SurfaceTile
{
    public int ID;
    public int Layer;
    public string SpriteName;
    public string Vertices;
}

[Serializable]
public class BuildingTile
{
    public int ID;
    public string SpriteName;
    public int GridWidth;
    public int GridHeight;
    public int Interactive;
}

