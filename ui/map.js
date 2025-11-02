let map, vectorSource, vectorLayer;

function init() {
  vectorSource = new ol.source.Vector({ features: [] });

  // Make the point big and obvious
  const pointStyle = new ol.style.Style({
    image: new ol.style.Circle({
      radius: 8,
      fill: new ol.style.Fill({ color: 'rgba(0,120,255,0.9)' }),
      stroke: new ol.style.Stroke({ color: 'white', width: 2 })
    })
  });

  vectorLayer = new ol.layer.Vector({ source: vectorSource, style: pointStyle });

  map = new ol.Map({
    target: 'map',
    layers: [
      // Basemap may fail on some networks; that's okay for now.
      new ol.layer.Tile({ source: new ol.source.OSM() }),
      vectorLayer
    ],
    view: new ol.View({
      center: ol.proj.fromLonLat([94.1088, 25.6754]),
      zoom: 10
    })
  });

  // Add one demo point
  const feature = new ol.Feature({
    geometry: new ol.geom.Point(ol.proj.fromLonLat([94.1088, 25.6754])),
    name: 'Demo point'
  });
  vectorSource.addFeature(feature);

  // Allow dragging the point
  const modify = new ol.interaction.Modify({ source: vectorSource });
  map.addInteraction(modify);
}

document.addEventListener('DOMContentLoaded', init);
