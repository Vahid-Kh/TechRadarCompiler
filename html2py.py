
"""

Property of Danfoss A/S  - RAC Tech Center
For support and code modifications contact:

    -Stefano Menengello (Python compiler support)
    -Vahid Khorshidi (HTML code & EXE Code generators support)

For executable file creation (".exe) do as follows:

    1- In Python terminal type "pip install pyinstaller"
    2- In Python terminal type "pip install pandas"
    3- In Python terminal type "pip install xlrd"
    4- Type in terminal "pyinstaller --onefile -w html2py.py"
    5- Put the html2py.exe file together with the TechRadarInputs.xlsx in the same directory
    6- Run html2py.exe and enjoy.

"""

"""

Save the excel file with input data in a local directory and paste it on filename variable
Specify the name of the radar
Specify the version of the radar

The radar will open automatically on your predefined browser and it will save itself as a html file in the same path
where this .py file is stored.

If you call two radars with the same name, the most recent will overwrite the older"""


import pandas as pd
import webbrowser

######## Inputs ############




# filename = r'System Protection/TechRadarInputs.xlsx'
# filename = r'System Control/TechRadarInputs.xlsx'
filename = r'Business unit 1/TechRadarInputs.xlsx'

# version = "V2.0"

####### end inputs ################


#############
product = pd.read_excel(filename, sheet_name="Techs", index_col= 0)
product = product.iloc[0][0]

techs = pd.read_excel(filename, sheet_name="Techs", index_col= 0, skiprows=2)

categories = pd.read_excel(filename, sheet_name="Categories", )

cat_list = []
for i in categories.index:
    cat_list.append(i)

cat_num = len(categories)


TechRadarName = r"" + str(product) + '\\RAC Tech Radar ' + product + ".html"
print(TechRadarName)
file = open(TechRadarName,'w', encoding='utf-8')

introA = """
<html>

<head>

  <meta http-equiv="Content-type" content="text/html; charset=utf-8">



  <title>    RAC Tech Radar """ + product + """</title>


  <script src="https://d3js.org/d3.v4.min.js"></script>
  <script>



    function radar_visualization(config) {

      // custom random number generator, to make random sequence reproducible
      // source: https://stackoverflow.com/questions/521295
      var seed = 42;
      function random() {
        var x = Math.sin(seed++) * 10000;
        return x - Math.floor(x);
      }

      function random_between(min, max) {
        return min + random() * (max - min);
      }

      function normal_between(min, max) {
        return min + (random() + random()) * 0.5 * (max - min);
      }

      // radial_min / radial_max are multiples of PI
"""

###### Sections #########

introB_3 = """
      const quadrants = [
        { radial_min: -1, radial_max: 0.5, factor_x: -1, factor_y: -1 },
        { radial_min: -0.3, radial_max: 0.3, factor_x: 1, factor_y: 1 },
        { radial_min: 0.5, radial_max: 1, factor_x: -1, factor_y: 1 },
      ];
"""

introB_4 = """
      const quadrants = [
        { radial_min: 0, radial_max: 0.5, factor_x: 1, factor_y: 1 },
        { radial_min: 0.5, radial_max: 1, factor_x: -1, factor_y: 1 },
        { radial_min: -1, radial_max: -0.5, factor_x: -1, factor_y: -1 },
        { radial_min: -0.5, radial_max: 0, factor_x: 1, factor_y: -1 }
      ];
"""

introB_5 = """
      const quadrants = 
      [
        { radial_min: -0.75, radial_max: -0.4, factor_x: -1, factor_y: -1 }, // 1
        { radial_min: 0.8, radial_max: 0.95, factor_x: -1, factor_y: 1 },  //  2
        { radial_min: 0.5, radial_max: 0.8, factor_x: -1, factor_y: 1 },  //  3
        { radial_min: -0.01, radial_max: -0.3, factor_x: 1, factor_y: -1 },  // 4
        { radial_min: 0, radial_max: 0.4, factor_x: 1, factor_y: 1 }, // 5
      
      ];
"""

introB_6 = """
      const quadrants = 
      [
        { radial_min: 0.0, radial_max: 0.99, factor_x: -1, factor_y: -1 }, 
        { radial_min: 0.0, radial_max: 0.2, factor_x: -1, factor_y: 1 },
        { radial_min: 0.01, radial_max: 0.39, factor_x: -1, factor_y: 1 },
        { radial_min: -0.7, radial_max: -0.33, factor_x: 1, factor_y: -1 },
        { radial_min: 0.2, radial_max: -0.02, factor_x: 1, factor_y: -1 },
        { radial_min: 0, radial_max: 0.23, factor_x: 1, factor_y: 1},  
      ];
"""

introB_7 = """
const quadrants = 
[
  { radial_min: -0.9, radial_max: -0.6, factor_x: -1, factor_y: -1 }, // 1
  { radial_min: -0.9, radial_max: -1, factor_x: -1, factor_y: -1 },  //  2
  { radial_min: 0.5, radial_max: 0.8, factor_x: -1, factor_y: 1 },  //  3
  { radial_min: 0.4, radial_max: 0.6, factor_x: 1, factor_y: 1 },  // 4
  { radial_min: -0.325, radial_max: -0.622, factor_x: 1, factor_y: -1 }, // 5
  { radial_min: -0.01, radial_max: -0.240, factor_x: 1, factor_y: -1 },  // 6
  { radial_min: 0.145, radial_max: 0.147, factor_x: 1, factor_y: 1}  // 7
];"""

introB_8 = """
const quadrants = 
[
  { radial_min: -0.73, radial_max: -0.52, factor_x: -1, factor_y: -1 }, 
  { radial_min: -0.98, radial_max: -0.77, factor_x: -1, factor_y: -1 },
  { radial_min: 0.77, radial_max: 0.98, factor_x: -1, factor_y: 1 },
  { radial_min: 0.52, radial_max: 0.73, factor_x: -1, factor_y: 1 },
  { radial_min: 0.27, radial_max: 0.48, factor_x: 1, factor_y: 1 }, 
  { radial_min: -0.05, radial_max: 0.23, factor_x: 1, factor_y: 1},
  { radial_min: -0.2, radial_max: -0.02, factor_x: 1, factor_y: -1 },
  { radial_min: -0.48, radial_max: -0.27, factor_x: 1, factor_y: -1 },

];
"""

introB = [introB_3, introB_4, introB_5, introB_6, introB_7, introB_8]

########################


introC= """
      const rings = [
        { radius: 130 },
        { radius: 220 },
        { radius: 310 },
        { radius: 410 }
      ];

      const title_offset =
      { x: -450, y: -500 };

      const footer_offset =
      { x: -120, y: 495 };
"""

########### Sections ################

introD_3 = """
     const legend_offset = [
       { x: -950, y: -310 },
       { x:  510 , y: -310 },
       { x: -950, y: 90 }
       
     ];

     const section_offset = [
       { x: -375, y: -310 },
       { x: -375, y: 440 },
       { x: 420, y: -30 }          
     ];
     
     
"""

introD_4 = """
      const legend_offset = [
        { x:  510, y: 90 },
        { x: -950, y: 90 },
        { x: -950, y: -310 },
        { x:  510, y: -310 }
      ];
      const section_offset = [
        { x: -430, y: -270 },
        { x: -480, y: 350 },
        { x: 300, y: -270 },
        { x: 300, y: 350 },
        ];
"""

introD_5 = """
      const legend_offset = [
        { x: -950, y: -350 },
        { x: -950, y: 30 },
        { x: -950, y: 260 },
        { x:  510, y: -310 },
        { x:  510, y: 100 },
      
      
      ];
      
      const section_offset = [
        { x: -305, y: -340 },
        { x: -505, y: -0 },
        { x: -205, y: 480 },
        { x:  250, y: -270 },
        { x:  280, y: 250 },
      ];
"""

introD_6 = """
      const legend_offset = [
        { x: -950, y: -450 },
        { x: -950, y: -120 },
        { x: -950, y: 190 },
      
        { x: 550, y: -450 },
        { x: 550, y: -120  },
        { x: 550, y: 190 },
      
        
      ];
      
      
      const section_offset = [
        { x: -550, y: -170 },
        { x: -550, y: 280 },
        { x: -40, y: 485},
      
        { x: 360, y: 280 },
        { x: 360, y: -170 },
        { x: -40, y: -380},
      
        
];
"""

introD_7 = """


const legend_offset = [
  { x: -950,  y: -450 },
  { x: -950,  y: -60  },
  { x: -950,  y:  160 },
  { x: -950,  y:  420 },
  { x:  550,  y: -450 },
  { x:  550,  y: -120 },
  { x:  550,  y:  120 }

];


const section_offset = [
  { x: -305,  y: -310 },
  { x: -455,  y:  -20 },
  { x: -385,  y:  320 },
  { x:    0,  y:  460 },
  { x:  350,  y:  300 },
  { x:  350,  y: -200 },
  { x:  100,  y: -300 }
];

"""

introD_8 = """


const legend_offset = [
  { x: -900,  y: -450  },
  { x: -755,  y: -450  },
  { x: -900,  y: 60    },
  { x: -755,  y: 100   },
  { x:  480,  y: 60    },
  { x:  640,  y: 100   },
  { x:  480,  y: -450  },
  { x:  640,  y: -320  },

  
];
 // /////////////////////////////////////////////////////////////////////

const section_offset = [
  { x:  -370, y: -350 },
  { x: -580, y: -120 },
  { x: -580, y: 230 },
  { x:  -300, y: 460 },
  { x: 120, y: 460 },
  { x: 290, y: 280 },
  { x:  350, y: -150 },
  { x: 120, y: -350 }
];

"""


introD = [introD_3, introD_4, introD_5, introD_6, introD_7, introD_8]
################################

introE = """
      function polar(cartesian) {
        var x = cartesian.x;
        var y = cartesian.y;
        return {
          t: Math.atan2(y, x),
          r: Math.sqrt(x * x + y * y)
        }
      }

      function cartesian(polar) {
        return {
          x: polar.r * Math.cos(polar.t),
          y: polar.r * Math.sin(polar.t)
        }
      }

      function bounded_interval(value, min, max) {
        var low = Math.min(min, max);
        var high = Math.max(min, max);
        return Math.min(Math.max(value, low), high);
      }

      function bounded_ring(polar, r_min, r_max) {
        return {
          t: polar.t,
          r: bounded_interval(polar.r, r_min, r_max)
        }
      }

      function bounded_box(point, min, max) {
        return {
          x: bounded_interval(point.x, min.x, max.x),
          y: bounded_interval(point.y, min.y, max.y)
        }
      }

      function segment(quadrant, ring) {
        var polar_min = {
          t: quadrants[quadrant].radial_min * Math.PI,
          r: ring === 0 ? 30 : rings[ring - 1].radius
        };
        var polar_max = {
          t: quadrants[quadrant].radial_max * Math.PI,
          r: rings[ring].radius
        };
        var cartesian_min = {
          x: 15 * quadrants[quadrant].factor_x,
          y: 15 * quadrants[quadrant].factor_y
        };
        var cartesian_max = {
          x: rings[3].radius * quadrants[quadrant].factor_x,
          y: rings[3].radius * quadrants[quadrant].factor_y
        };
        return {
          clipx: function (d) {
            var c = bounded_box(d, cartesian_min, cartesian_max);
            var p = bounded_ring(polar(c), polar_min.r + 15, polar_max.r - 15);
            d.x = cartesian(p).x; // adjust data too!
            return d.x;
          },
          clipy: function (d) {
            var c = bounded_box(d, cartesian_min, cartesian_max);
            var p = bounded_ring(polar(c), polar_min.r + 15, polar_max.r - 15);
            d.y = cartesian(p).y; // adjust data too!
            return d.y;
          },
          random: function () {
            return cartesian({
              t: random_between(polar_min.t, polar_max.t),
              r: normal_between(polar_min.r, polar_max.r)
            });
          }
        }
      }

      // position each entry randomly in its segment
      for (var i = 0; i < config.entries.length; i++) {
        var entry = config.entries[i];
        entry.segment = segment(entry.quadrant, entry.ring);
        var point = entry.segment.random();
        entry.x = point.x;
        entry.y = point.y;
        entry.color = entry.active || config.print_layout ?
          config.flame[entry.flame].color : config.colors.inactive;
      }

      // partition entries according to segments
      var segmented = new Array(10);
      for (var quadrant = 0; quadrant < 10; quadrant++) {
        segmented[quadrant] = new Array(4);
        for (var ring = 0; ring < 4; ring++) {
          segmented[quadrant][ring] = [];
        }
      }
      for (var i = 0; i < config.entries.length; i++) {
        var entry = config.entries[i];
        segmented[entry.quadrant][entry.ring].push(entry);
      }

      // assign unique sequential id to each entry
      var id = 1;
      for (var quadrant of """ + str(cat_list) + """) {
        for (var ring = 0; ring < 4; ring++) {
          var entries = segmented[quadrant][ring];
          entries.sort(function (a, b) { return a.label.localeCompare(b.label); })
          for (var i = 0; i < entries.length; i++) {
            entries[i].id = "" + id++;
          }
        }
      }

      function translate(x, y) {
        return "translate(" + x + "," + y + ")";
      }

      function viewbox(quadrant) {
        return [
          Math.max(0, quadrants[quadrant].factor_x * 410) - 420,
          Math.max(0, quadrants[quadrant].factor_y * 410) - 420,
          440,
          440
        ].join(" ");
      }

      var svg = d3.select("svg#" + config.svg_id)
        .style("background-color", config.colors.background)
        .attr("width", config.width)
        .attr("height", config.height);

      var radar = svg.append("g");
      if ("zoomed_quadrant" in config) {
        svg.attr("viewBox", viewbox(config.zoomed_quadrant));
      } else {
        radar.attr("transform", translate(config.width / 2, config.height / 2));
      }

      var grid = radar.append("g");
"""

############# Sections ################
introF_3 = """
      // draw grid lines
      grid.append("line")
        .attr("x1", -0).attr("y1", 0)
        .attr("x2", -400).attr("y2", 0)
        .style("stroke", config.colors.grid)
        .style("stroke-width", 2);
      grid.append("line")
        .attr("x1", 0).attr("y1", 0)
        .attr("x2", 205).attr("y2", 355.07)
        .style("stroke", config.colors.grid)
        .style("stroke-width", 2);
      grid.append("line")
        .attr("x1", 0).attr("y1", 0)
        .attr("x2", 205).attr("y2", -355.07)
        .style("stroke", config.colors.grid)
        .style("stroke-width", 2);
      
"""

introF_4 = """
      // draw grid lines
      grid.append("line")
        .attr("x1", 0).attr("y1", -410)
        .attr("x2", 0).attr("y2", 410)
        .style("stroke", config.colors.grid)
        .style("stroke-width", 2);
      grid.append("line")
        .attr("x1", -410).attr("y1", 0)
        .attr("x2", 410).attr("y2", 0)
        .style("stroke", config.colors.grid)
        .style("stroke-width", 2);
"""

introF_5 = """
      // draw grid lines // to change number of categories
      grid.append("line")
      .attr("x1", 0).attr("y1", 0)
      .attr("x2", 410).attr("y2", 0)
      .style("stroke", config.colors.grid)
      .style("stroke-width", 2);
      grid.append("line")
        .attr("x1", 0).attr("y1", 0)
        .attr("x2", 126.7).attr("y2", 389.93)
        .style("stroke", config.colors.grid)
        .style("stroke-width", 2);
      grid.append("line")
      .attr("x1", 0).attr("y1", 0)
      .attr("x2", -331).attr("y2", 240)
      .style("stroke", config.colors.grid)
      .style("stroke-width", 2);
      
      grid.append("line")
      .attr("x1", 0).attr("y1", 0)
      .attr("x2", -331.7).attr("y2", -240.99)
      .style("stroke", config.colors.grid)
      .style("stroke-width", 2);
      
      grid.append("line")
      .attr("x1", 0).attr("y1", 0)
      .attr("x2", 126.7).attr("y2", -389.82)
      .style("stroke", config.colors.grid)
      .style("stroke-width", 2);
"""

introF_6 = """
      // draw grid lines // to change number of categories
      
      grid.append("line")
        .attr("x1", -410).attr("y1", 0)
        .attr("x2", 410).attr("y2", 0)
        .style("stroke", config.colors.grid)
        .style("stroke-width", 2);
      grid.append("line")
        .attr("x1", -205).attr("y1", 355.07)
        .attr("x2", 205).attr("y2", -355.07)
        .style("stroke", config.colors.grid)
        .style("stroke-width", 2);
      grid.append("line")
        .attr("x1", 205).attr("y1", 355.07)
        .attr("x2", -205).attr("y2", -355.07)
        .style("stroke", config.colors.grid)
        .style("stroke-width", 2);
"""

introF_7 = """

// draw grid lines // to change number of categories
grid.append("line")
.attr("x1", 0).attr("y1", 0)
.attr("x2", 410).attr("y2", 0)
.style("stroke", config.colors.grid)
.style("stroke-width", 3);
grid.append("line")
  .attr("x1", 0).attr("y1", 0)
  .attr("x2", 255.6).attr("y2", 320.51)
  .style("stroke", config.colors.grid)
  .style("stroke-width", 3);
grid.append("line")
.attr("x1", 0).attr("y1", 0)
.attr("x2", -88).attr("y2", 400.4)
.style("stroke", config.colors.grid)
.style("stroke-width", 3);

grid.append("line")
.attr("x1", 0).attr("y1", 0)
.attr("x2", -367.2).attr("y2", 182.3)
.style("stroke", config.colors.grid)
.style("stroke-width", 3);

grid.append("line")
.attr("x1", 0).attr("y1", 0)
.attr("x2", -372.18).attr("y2", -171.97)
.style("stroke", config.colors.grid)
.style("stroke-width", 3);
grid.append("line")
.attr("x1", 0).attr("y1", 0)
.attr("x2", -99.18).attr("y2", -397.82)
.style("stroke", config.colors.grid)
.style("stroke-width", 3);

grid.append("line")
.attr("x1", 0).attr("y1", 0)
.attr("x2", 247.88).attr("y2", -326.57)
.style("stroke", config.colors.grid)
.style("stroke-width", 3);
grid.append("line")
.attr("x1", 0).attr("y1", 0)
.attr("x2", 255.6).attr("y2", 320.51)
.style("stroke", config.colors.grid)
.style("stroke-width", 3);

"""

introF_8 = """

// draw grid lines // to change number of categories
grid.append("line")
  .attr("x1", 0).attr("y1", -410)
  .attr("x2", 0).attr("y2", 410)
  .style("stroke", config.colors.grid)
  .style("stroke-width", 3);
grid.append("line")
  .attr("x1", -410).attr("y1", 0)
  .attr("x2", 410).attr("y2", 0)
  .style("stroke", config.colors.grid)
  .style("stroke-width", 3);
grid.append("line")
  .attr("x1", -290).attr("y1", 290)
  .attr("x2", 290).attr("y2", -290)
  .style("stroke", config.colors.grid)
  .style("stroke-width", 3);
grid.append("line")
  .attr("x1", 290).attr("y1", 290)
  .attr("x2", -290).attr("y2", -290)
  .style("stroke", config.colors.grid)
  .style("stroke-width", 3);


"""

introF = [introF_3, introF_4, introF_5, introF_6, introF_7, introF_8]

#############################################


introG = """
      // background color. Usage `.attr("filter", "url(#solid)")`
      // SOURCE: https://stackoverflow.com/a/31013492/2609980
      var defs = grid.append("defs");
      var filter = defs.append("filter")
        .attr("x", 0)
        .attr("y", 0)
        .attr("width", 1)
        .attr("height", 1)
        .attr("id", "solid");
      filter.append("feFlood")
        .attr("flood-color", "rgb(0, 0, 0, 0.8)");
      filter.append("feComposite")
        .attr("in", "SourceGraphic");

      // draw rings
      for (var i = 0; i < rings.length; i++) {
        grid.append("circle")
          .attr("cx", 0)
          .attr("cy", 0)
          .attr("r", rings[i].radius)
          .style("fill", "none")
          .style("stroke", config.colors.grid)
          .style("stroke-width", 2);
        if (config.print_layout) {
          grid.append("text")
            .text(config.rings[i].name)
            .attr("y", -rings[i].radius + 52)
            .attr("text-anchor", "middle")
            .style("fill", "#ab8484")
            .style("font-family", "Arial, Helvetica")
            .style("font-size", 42)
            .style("font-weight", "bold")
            .style("pointer-events", "none")
            .style("user-select", "none");
        }
      }

      function legend_transform(quadrant, ring, index = null) {
        var dx = ring < 12 ? 0 : 120;
        var dy = (index == null ? -16 : index * 10);
        if (ring === 0) {
          dy = dy ;
        } else if (ring === 1) {
          dy = dy + 10 + segmented[quadrant][ring - 1].length * 11;
        } else if (ring === 2) {
          dy = dy + 100 + segmented[quadrant][ring - 1].length * 11;
        } else if (ring === 3) {
          dy = dy + 140 + segmented[quadrant][ring - 1].length * 11;
        }

        return translate(
          legend_offset[quadrant].x + dx,
          legend_offset[quadrant].y + dy
        );
      }

      // draw title and legend (only in print layout)
      if (config.print_layout) {

        // title
        radar.append("text")
          .attr("transform", translate(title_offset.x, title_offset.y))
          .text(config.title)
          .style("font-family", "Arial, Helvetica")
          .style("font-size", "48");

        // footer
        radar.append("text")
          .attr("transform", translate(footer_offset.x-300, footer_offset.y))
          .text("Technology feasibility:     ★ High       ● Medium      ▼ Low")
          .attr("xml:space", "preserve")
          .style("fill", "#616161")
          .style("font-family", "Arial, Helvetica")
          .style("font-weight", "bold")
          .style("font-size", "30");
        
        // footer color
                radar.append("text")
          .attr("transform", translate(footer_offset.x-300, footer_offset.y+50))
          .text("Market desirability: ")
          .attr("xml:space", "preserve")
          .style("fill", "#616161")
          .style("font-family", "Arial, Helvetica")
          .style("font-weight", "bold")
          .style("font-size", "30");
                // footer color
        radar.append("text")
          .attr("transform", translate(footer_offset.x+100, footer_offset.y+50))
          .text("High")
          .attr("xml:space", "preserve")
          .style("fill", "#2CE71E")
          .style("font-family", "Arial, Helvetica")
          .style("font-weight", "bold")
          .style("font-size", "30");
                // footer color
        radar.append("text")
          .attr("transform", translate(footer_offset.x+250, footer_offset.y+50))
          .text(" Medium")
          .attr("xml:space", "preserve")
          .style("fill", "#FFA500")
          .style("font-family", "Arial, Helvetica")
          .style("font-weight", "bold")
          .style("font-size", "30");
                // footer color

        radar.append("text")
          .attr("transform", translate(footer_offset.x+450, footer_offset.y+50))
          .text(" Low ")
          .attr("xml:space", "preserve")
          .style("fill", "#bf0000")
          .style("font-family", "Arial, Helvetica")
          .style("font-weight", "bold")
          .style("font-size", "30");


        // legend
        var legend = radar.append("g");
        for (var quadrant = 0; quadrant < """ + str(cat_num) + """; quadrant++) {
          legend.append("text")
            .attr("transform", translate(
              legend_offset[quadrant].x,
              legend_offset[quadrant].y - 45
            ))
            .text(config.quadrants[quadrant].name)
            .style("font-family", "Arial, Helvetica")
            .style("fill", "#371f72")
            .style("font-weight", "bold")
            .style("font-size", "18");

            
            legend.append("text")
              .attr("transform", translate(
                section_offset[quadrant].x,
                section_offset[quadrant].y - 45
        ))
            .text(config.quadrants[quadrant].name)
            .style("font-family", "Arial, Helvetica")
            .style("fill", "#371f72")
            .style("font-weight", "bold")
            .style("font-size", "18");


          for (var ring = 0; ring < 4; ring++) {
            legend.append("text")
              .attr("transform", legend_transform(quadrant, ring))
              .text(config.rings[ring].name)
              .style("font-family", "Arial, Helvetica")
              .style("font-size", "13")
              .style("font-weight", "bold");
            legend.selectAll(".legend" + quadrant + ring)
              .data(segmented[quadrant][ring])
              .enter()

              .append("a")
              .attr("xlink:href", function (d, i) {
                return d.url ? "#" : null; // it's better to open new window so we provide # as url to stay on same page
              })
              .on("click", function (d) {
                // let's open new tab/window instead of redirecting off the radar
                window.open(d.url, "_blank");
              })

              .append("text")
              .attr("transform", function (d, i) { return legend_transform(quadrant, ring, i); })
              .attr("class", "legend" + quadrant + ring)
              .attr("id", function (d, i) { return "legendItem" + d.id; })
              .text(function (d, i) { return d.id + ". " + d.label; })
              .style("font-family", "Arial, Helvetica")
              .style("font-size", "12")
              .on("mouseover", function (d) { showBubble(d); highlightLegendItem(d); })
              .on("mouseout", function (d) { hideBubble(d); unhighlightLegendItem(d); });
          }
        }
      }

      // layer for entries
      var rink = radar.append("g")
        .attr("id", "rink");

      // rollover bubble (on top of everything else)
      var bubble = radar.append("g")
        .attr("id", "bubble")
        .attr("x", 0)
        .attr("y", 0)
        .style("opacity", 0)
        .style("pointer-events", "none")
        .style("user-select", "none");
      bubble.append("rect")
        .attr("rx", 4)
        .attr("ry", 4)
        .style("fill", "#333");
      bubble.append("text")
        .style("font-family", "sans-serif")
        .style("font-size", "20px")
        .style("fill", "#fff");
      bubble.append("path")
        .attr("d", "M 0,0 10,0 5,8 z")
        .style("fill", "#333");

      function showBubble(d) {
        if (d.active || config.print_layout) {
          var tooltip = d3.select("#bubble text")
            .text(d.label);
          var bbox = tooltip.node().getBBox();
          d3.select("#bubble")
            .attr("transform", translate(d.x - bbox.width / 2, d.y - 16))
            .style("opacity", 0.8);
          d3.select("#bubble rect")
            .attr("x", -5)
            .attr("y", -bbox.height)
            .attr("width", bbox.width + 10)
            .attr("height", bbox.height + 4);
          d3.select("#bubble path")
            .attr("transform", translate(bbox.width / 2 - 5, 3));
        }
      }

      function hideBubble(d) {
        var bubble = d3.select("#bubble")
          .attr("transform", translate(0, 0))
          .style("opacity", 0);
      }

      function highlightLegendItem(d) {
        var legendItem = document.getElementById("legendItem" + d.id);
        legendItem.setAttribute("filter", "url(#solid)");
        legendItem.setAttribute("fill", "white");
      }

      function unhighlightLegendItem(d) {
        var legendItem = document.getElementById("legendItem" + d.id);
        legendItem.removeAttribute("filter");
        legendItem.removeAttribute("fill");
      }

      // draw blips on radar
      var blips = rink.selectAll(".blip")
        .data(config.entries)
        .enter()
        .append("g")
        .attr("class", "blip")
        .attr("transform", function (d, i) { return legend_transform(d.quadrant, d.ring, i); })
        .on("mouseover", function (d) { showBubble(d); highlightLegendItem(d); })
        .on("mouseout", function (d) { hideBubble(d); unhighlightLegendItem(d); });

      // configure each blip
      blips.each(function (d) {
        var blip = d3.select(this);

        // blip link
        if (!config.print_layout && d.active && d.hasOwnProperty("link")) {
          blip = blip.append("a")
            .attr("xlink:href", d.link);
        }

        // blip shape // TO CHANGE BLIP SHAPE!!
        if (d.moved > 0) {
          blip.append("path")
            .attr("d", "M 0,-20 -11.755,16.18 19.02,-6.18 -19.02,-6.18 11.755,16.18 0,-20 z") // Highlight 20,10 -20,10 15,-20 0,30 30,-20 20,10 Star
            .style("fill", d.color);
        } else if (d.moved < 0) {
          blip.append("path")
          .attr("d", "M -16,-9 16,-9 0,13 z") // triangle pointing down
          .style("fill", d.color);
        } else {
          blip.append("circle")
            .attr("r", 12)
            .attr("fill", d.color);
        }

        // blip text
        if (d.active || config.print_layout) {
          var blip_text = config.print_layout ? d.id : d.label.match(/[a-z]/i);
          blip.append("text")
            .text(blip_text)
            .attr("y", 3)
            .attr("text-anchor", "middle")
            .style("fill", "#fff")
            .style("font-family", "Arial, Helvetica")
            .style("font-size", function (d) { return blip_text.length > 2 ? "11" : "12"; })
            .style("pointer-events", "none")
            .style("user-select", "none");
        }
      });

      // make sure that blips stay inside their segment
      function ticked() {
        blips.attr("transform", function (d) {
          return translate(d.segment.clipx(d), d.segment.clipy(d));
        })
      }

      // distribute blips, while avoiding collisions
      d3.forceSimulation()
        .nodes(config.entries)
        .velocityDecay(0.19) // magic number (found by experimentation)
        .force("collision", d3.forceCollide().radius(12).strength(0.85))
        .on("tick", ticked);
    }

  </script>

  <style>
    body {
      font-family: 'Source Sans Pro', arial, helvetica, sans-serif;
      padding-bottom: 50px;
    }

    h3 {
      margin-top: 50px;
    }

    li {
      margin: 25px 50px 0 0;
    }

    table {
      width: 2000px;
      margin: 0 50px 0 50px;
    }

    td {
      width: 50%;
      vertical-align: top;
      padding-right: 60px;
    }
  </style>
"""

########### Write different number of sections ############

k = cat_num - 3

file.write(introA)
file.write(introB[k])
file.write(introC)
file.write(introD[k])
file.write(introE)
file.write(introF[k])
file.write(introG)




entries_intro1 = """
</head>

<body>

  <svg id="radar"></svg>

  <script>
    radar_visualization({
      svg_id: "radar",
      width: 2000,
      height: 1200,
      colors: {
        background: "#fff",
        grid: "#545454",
        inactive: "#545454"
      },
      title: "Product Tech Radar - """ + product + """",
      
      quadrants: [
      """
file.write(entries_intro1)


for i in categories.index:
        entries_cat = """
        { name: " """ + categories.loc[i, 'Categories'] + """ " },
        """

        file.write(entries_cat)

entries_intro2 = """
      ],
      rings: [
        { name: "ADOPT", color: "#93c47d" },
        { name: "TRIAL", color: "#93d2c2" },
        { name: "ASSESS", color: "#fbdb84" },
        { name: "HOLD", color: "#efafa9" }
      ],
      
      
      flame: [
        {color: "#2CE71E" },
        {color: "#FFA500" },
        {color: "#bf0000" }
      ],
      
      print_layout: true,
      //ENTRIES
      // 0,1,2,3 (counting clockwise, starting from bottom right)
      // 0,1,2,3 (starting from inside)
      // 1 = High (Filled star)
      //  0 = Medium (circle)
      //  -1 =Low (triangle pointing down)

      entries: [
      """

file.write(entries_intro2)

import math
for i in techs.index:
    print(i, ' ',techs.loc[i, 'Technology title'])
    if techs.loc[i, 'One Pager'] != 'no':
        onepager = techs.loc[i, 'One Pager']

    else:
        onepager = "https://danfoss.sharepoint.com/:p:/s/ProductTechnologyStrategy/EdHOjH4yPJ5BsUOlmHI__wkBa6t5Cm90OBIrEhDjp-lucA?e=yGWtAm"

    stringa = """
            {
                quadrant: """ + str(techs.loc[i, 'Category']) + """,
                ring: """ + str(techs.loc[i, 'Effort level']) + """,
                label: " """ + techs.loc[i, 'Technology title'] + """ ",
                moved: """ + str(techs.loc[i, 'Importance']) + """,
                flame: """ + str(techs.loc[i, 'Color']) + """,
                url: " """ + onepager + """ "
                
            },
    """
    file.write(stringa)

entries_concl = """
      ]
      //ENTRIES
    });
  </script>


</body>
"""

file.write(entries_concl)

conclusion = """
</html>
    <h1> 
        <p style="text-align:center;">     
             <strong> Got ideas... Give it a chance to come alive:   <a href="https://innovationforall.danfoss.com/ideaboxes/608914758abe520e62ba4345">RAC Innovation for All</a>.</strong>
        </p>
    </h1>
    
    <body>
<h4>
      <p style="text-align:center;">
More information can be found here: 
    <a href="https://danfoss.sharepoint.com/:p:/s/RACTechRoadmap/EcLf1NRBCRNCmcKyXIZOlzsB5hahz9_f1yrqZm81or4hcg?email=vk%40danfoss.com&e=gpqTXO">Product Tech Radar</a>.
    <a href="https://danfoss.sharepoint.com/:p:/s/RACTechRoadmap/EZ6TT6sdaspFsYkgpti4ztMBiTH99_3OAkbS6N7fIcFOFQ?email=vk%40danfoss.com&e=NUjPO1">Product Tech Strategy</a>.
</h4>
</p>
<h4>
      <p style="text-align:center;">
Product Tech Strategy: 
    <a href="https://miro.com/app/board/o9J_lUSZ8s8=/">Expansion Function</a>, 
    <a href="https://miro.com/app/board/o9J_lU98UmU=/">System Control</a> & 
    <a href="https://miro.com/app/board/o9J_lUw-pNg=/">System Protection</a>.
</h4>
</p>

<h4>
<p>
      <p style="text-align:center;">"Technology Feasibility" & "Market Desirability" overview: <a href="https://miro.com/app/board/o9J_lRN4N9k=/?moveToWidget=3074457361877266308&cot=14">RAC Tech Radar Evaluation Matrix</a>.</p>

</h4>
<h4>
<p>

      <p style="text-align:center;">                           Disclaimer : 
This tech radar is a first draft. "Technology Feasibility" & "Market Desirability"  are somewhat randomly assigned; a selected committee is required to judge the real positions.
</p>
</p>
</h4>
</body>

<table>
  <tr>
    <td>



      <h3>What is the Tech Radar?</h3>

      <p>
        
        Tech is changing at a breakneck speed. Tech radar is a tool to map out a tech strategy that will keep us ahead of the curve.
        The Tech Radar contains a thoughtfully selected list of relevant technologies, complemented by an assessment result, called <em>ring effort level</em>.
        We use four rings with the following semantics
        :
      </p>

      <ul>
        <li><strong>Adopt: On the NPD roadmap</strong>
        <br>Tech is currently in the new product development pipeline
        <br>Selected committee of PM, R&D, sales & application experts deciding on which technology projects to proceed into the NPD pipeline. </li>
        <li><strong>Trial: Tech exploration</strong>
        <br>This is a prototype and test technology project 
        <br>There is pilot running project either internally or with partner/supplier</li>
        <li><strong>Assess: Tech assessment </strong>;
        <br>There is tech assessment project either internally or with partner/supplier
        <br>University/research institute collaborations or with partners, side tasks of specialists - limited internal resource allocation. </li>
        <li><strong>Hold: Tech monitoring</strong> 
        <br>Monitored on IP and research level with low internal resource allocation
        <br>This technology has very low maturity or outdated/deprecated.
        <br>There is high uncertainty of using this technology in new projects.</li>
      </ul>
      </p>
      



    </td>
    <td>


      <h3>What is the value?</h3>

The technology radar is a systematic & unbiased process that aims to fill the pre-NPD pipeline with technology innovations following strategies aligned with business.

<li>Scouting & development of technologies to support RAC technology strategy 
<li>Monitoring tool of interesting innovative technologies at an early stage
<li>Strike a balance between Market Pull vs. Tech Push 
<li>Periodical meetings to bridge technology and business by strategic theme oriented approach 
<li>Democratization of scouted technologies


<h3>Background</h3>
<p> The Tech Radar is a tool to inspire and support engineering teams to pick the best technologies for new projects;
        it provides a platform to share knowledge and experience in technologies, to reflect on technology decisions and
         continuously evolve our technology landscape. Based on the <a href="https://www.thoughtworks.com/radar">pioneering work of ThoughtWorks</a>.  </p>

 <li>  We would like to here your feedbacks/comments, please share it with us at: vk@danfoss.com 
    </td>

  </tr>

</table>


</body>


</html>




"""

file.write(conclusion)

file.close()
webbrowser.open_new_tab(TechRadarName)




