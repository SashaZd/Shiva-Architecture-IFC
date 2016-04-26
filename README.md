Shiva-Architecture-IFC
======================

Custom IFC Parser &amp; Algorithm generated for a PhD Architecture Project @Georgia Institute of Technology. This code enabled the Semantic Knowledge Based Rule Set described in the [PhD Dissertation](https://smartech.gatech.edu/handle/1853/53496) of Georgia Tech's student [Shiva Aram](https://www.linkedin.com/in/shiva-aram-97179015) to be automated.

The code is written in Python to automatically write the results of precast concrete slab modularization into the enriched IFC file as well as Excel sheets which is the form that cost estimators usually use for QTO and CE activities. The following is the information that are provided for users in the output files:

* An enriched IFC file
  * Equal to the number of the slab pieces that the algorithm devises for each slab to be segmented into, IfcSlab entities are created and added to the end of IFC file. Examples of these added entities can be seen in Figure 5.11. Slab width, floor level, span length and number of strands and rebars used in its stem design were added to Name, Description, ObjectType and Tag attributes of the entities as seen in Figure 5.11.
* Excel tables: 
  * In the first table the slab piece is organized per floor level. The tables and their information items were designed based on actual QTO tables that were collected from precast concrete companies. The code finds all the slabs with equal width and span length that are in the same floor level and writes their size and concrete and reinforcement quantity information in one row. Quantities were first provided per piece and then total quantity of same size slabs in each level is provided. Total concrete volume was calculated multiplying the DT slab profile area by its span and by number of DT pieces in each floor, where DT slab profile area was extracted from PCI Handbook. 
  * In the second table the same information was provided for DTs of similar size in the whole project as well as total linear feet, volume and weight of concrete used in the whole project.

The algorithm was designed in a way that it minimized the number of slab pieces while adhering to user preferences and limitations. The algorithm was designed by [Shiva Aram](https://www.linkedin.com/in/shiva-aram-97179015) and it's subsequent use and a more detailed description has been defined in her [PhD Dissertation](https://smartech.gatech.edu/handle/1853/53496)
