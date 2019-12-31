import rtree
import fiona
import shapely
import pyproj
import pandas as pd
import geopandas as gpd
from requests import Request
from owslib.wfs import WebFeatureService
import matplotlib
from matplotlib import pyplot as plt
matplotlib.use('TkAgg')
import mapclassify
import pysal
import urbanaccess
import pandana
import dill
import tk

########################################################################################################################
# import WFS
url = "https://geoservices.informatievlaanderen.be/overdrachtdiensten/VRBG2019/wfs?"
wfs = WebFeatureService(url=url)

# define layer ~ get capabilities(list layers)
layers = list(wfs.contents)
print(layers)
chosen_layer = "VRBG2019:Refgem"

# import chosen layer
params = dict(service='WFS', version="1.0.0", request='GetFeature',
      typeName=chosen_layer, outputFormat='json')
q = Request('GET', url, params=params).prepare().url

# assign layer as geopandas object
wfs_layer = gpd.read_file(q)

########################################################################################################################
# import shapefile
matplotlib.use('TkAgg')  # to avoid matplotlib runs in agg backend which cannot show plots
                         # (due to pip installs of some libraries)

data = gpd.read_file("D:\KUL\Master\stage\GeoSparc\geopandas\data\Refgem.shp")
data2 = gpd.read_file("D:\KUL\Master\stage\GeoSparc\geopandas\data\Refprv.shp")

print(data.columns)   # summarizes columns
print(data.head())    # gives first 5 rows

# plot
data.plot()           # plot data geometry

########################################################################################################################
# import csv as pandas dataframe
cities = pd.read_csv("D:/KUL/Master/stage/GeoSparc/geopandas/data/belgian-cities.csv", sep=',')
print(cities.head())

# convert pd dataframe to gpd dataframe
from shapely.geometry import Point
geom = cities.apply(lambda x: Point([x['lng'], x['lat']]), axis=1)  # form geometry
cities_gpd = gpd.GeoDataFrame(cities, geometry=geom, crs={'init': 'epsg:4326'})

# convert crs
print(cities_gpd.crs)
cities_gpd = cities_gpd.to_crs(data.crs)
print(cities_gpd.crs)
print(data2.crs)

#plot
matplotlib.use('TkAgg')
base = data2.plot(color='none', edgecolor='black')
cities_gpd.plot(ax=base)
plt.show()

########################################################################################################################
# plotting

# plot data geometry with variable for colormap
data.plot(column='OPPERVL', cmap='rainbow', legend=True, legend_kwds={'label': "Oppervlakte",
                                                                      'orientation': "horizontal"})

# scheme option forms categorical division, no legend_kwds possible
data.plot(column='OPPERVL', scheme='equal_interval', cmap='winter', edgecolor='black', legend=True)
# https://matplotlib.org/3.1.0/tutorials/colors/colormaps.html for more cmap schemes

# plot scheme provided by mapclassify
# ‘box_plot’, ‘equal_interval’, ‘fisher_jenks’, ‘fisher_jenks_sampled’, ‘headtail_breaks’, ‘jenks_caspall’,
# ‘jenks_caspall_forced’, ‘jenks_caspall_sampled’, ‘max_p_classifier’, ‘maximum_breaks’, ‘natural_breaks’,
# ‘quantiles’, ‘percentiles’, ‘std_mean’ or ‘user_defined’

plt.show()
########################################################################################################################
# multiple layers
# transformation of coordinate reference system (CRS)
print(data.crs)
print(data2.crs)
data3 = data2.to_crs(data.crs)

base = data.plot(column='OPPERVL', cmap='winter')
data3.plot(ax=base, color='none', edgecolor='black')

plt.show()

########################################################################################################################
# filter 1
Gent = data[data['NAAM'] == 'Gent']
print(Gent)
print(Gent.values)
print(Gent.describe())

# filter 2
Grote_gem = data[data['OPPERVL'] > 100000000]

# plot
base = Grote_gem.plot(color='red', edgecolor='none')
Gent.plot(ax=base, color='green', edgecolor='none')
data.plot(ax=base, color='none', edgecolor='grey')
data2.plot(ax=base, color='none', edgecolor='black')
plt.show()

Gent = data[data['NAAM'] == 'Gent']

invloedsfeer = gpd.GeoDataFrame({'geometry': Gent.buffer(distance=20000)})  # overlay works only for GeoDataFrames
from geopandas.tools import overlay
invloed_bel = overlay(invloedsfeer, data, how='intersection')

base = invloedsfeer.plot(color='yellow')
invloed_bel.plot(ax=base, color='red')
Gent.plot(ax=base, color='none', edgecolor='black')
data2.plot(ax=base, color='none', edgecolor='black')
plt.show()
