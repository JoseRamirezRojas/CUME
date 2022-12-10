<template>

<div class="row">
  <div class="column" style="background-color: #e3f2fd">
    <MDBListGroup  id="inner" >
      <MDBListGroupItem tag="button" @click="navega('/protocolos-info')" action >Protocolos</MDBListGroupItem>
      <MDBListGroupItem tag="button" @click="navega('/protocolo-fisicoquimica')" action> Calidad fisicoquímica </MDBListGroupItem>
      <MDBListGroupItem tag="button" @click="navega('/protocolo-hyqi')" action > HYQI </MDBListGroupItem>
      <MDBListGroupItem tag="button" @click="navega('/protocolo-deqi')" action>DEQI</MDBListGroupItem>
      <MDBListGroupItem tag="button" @click="navega('/protocolo-meqi')" action >MEQI</MDBListGroupItem>
      <MDBListGroupItem tag="button" action active>BEMEQI </MDBListGroupItem>
      <MDBListGroupItem tag="button" @click="navega('/cuenca-jurisdiccion')" action >Material de trabajo</MDBListGroupItem>
    </MDBListGroup>

  </div>

  <div class="column-2" style="background-color:#F8F8F8;">
    <MDBTabs v-model="activeTabId">
      <!-- Tabs navs -->
      <MDBTabNav pills justify tabsClasses="mb-3">
        <MDBTabItem tabId="info-bemeqi" href="info-bemeqi"      >Información del protocolo</MDBTabItem>
        <MDBTabItem tabId="calculo-bemeqi" href="calculo-bemeqi">Cálculo de índice BEMEQI</MDBTabItem>
      </MDBTabNav>
      <!-- Tabs navs -->
      <!-- Tabs content -->
      <MDBTabContent>
        <!-- info protocolo-->
        <MDBTabPane tabId="info-bemeqi">
          <h1 class="display-6">Macroinvertebrados </h1>
          <p align="justify" > Los macroinvertebrados lo integran un conjunto de grupos de taxa con diferente historia evolutiva que se han adaptado a las mismas condiciones que se encuentran en los ríos. La mayoría de ellos son bentónicos, por lo que dependen del tipo, tamaño y cantidad de sustrato disponibles en el cuerpo de agua donde habitan. </p>
          <img src="../assets/bemeqi_introduct.png"  class="img-fluid  hover-shadow" alt="bemeqi" style="float:left;  width: 197px; margin-right: 15px" />
          <p align="justify" > Además dentro de la gama de indicadores biológicos, los macroinvertebrados son quizás el grupo más usado para el monitoreo de los cuerpos de agua.</p>
          <p align="justify" > Son fundamentales en el mantenimiento de la cadena alimenticia en el concepto del contínuo del río. Y debido a la diversidad evolutiva de los distintos taxa, es seguro que al menos algunos de los organismos que conforman los ensambles van a reaccionar a cambios de origen natural o antropogénico.</p>
          <p align = "justify"> Debido a la naturaleza heterogénea de los ríos, se recomienda establecer un transecto de 10 metros de longitud en el tramo de río a evaluar, donde sea posible establecer diferentes zonas que abarcan tipos de sustratos y regímenes de velocidad y profundidad contrastantes. Se deberán seguir las mismas recomendaciones y de preferencia, elegir los mismos sitios donde se realizó la colecta de macroalgas y diatomeas. </p>
          <h1 class="display-6">Colecta </h1>
          <p align = "justify"> Para colectarlos se pueden seguir varios métodos entre los que resalta la <b>técnica del pateo</b>, la cual requiere de la remoción mecánica del sustrato en dirección opuesta del flujo de agua y cubrir los diferentes hábitats que han sido detectados visualmente.</p>
          <img src="../assets/pateo.png"  class="img-fluid  hover-shadow" alt="pateo" style=" width: 460px; " /> <br> <br> 
          <img src="../assets/colecta_bemeqi.png"  class="img-fluid  hover-shadow" alt="transecto" style="float:right; width: 141px; margin-left: 20px" /> <br>
          <p align="justify" >  Esta colecta se sugiere con una red tipo “D” que tenga una apertura de malla de 250 micrómetros. Se recomienda que el pateo sea por un periodo aproximado de cinco minutos para evitar el sobremuestreo. Este criterio también será avalado por la cantidad de individuos que se colecten y cuando ya no se visualicen nuevos taxones.</p>
          <p align="justify"> La colecta de los diferentes hábitats debe ser depositada en una charola blanca con la finalidad de tener una muestra compuesta del transecto a evaluar. Después se procede, con la ayuda de pinzas entomológicas, a la selección de los organismos colectados para que sean depositados en un frasco de colecta. </p>
          <img src="../assets/red_colecta_bemeqi.png"  class="img-fluid  hover-shadow" alt="bemeqi" style="float:left;  width: 246px; margin-right: 15px" />
          <p align="justify"> Debido a la diversidad de hábitats posibles donde se encuentran los macroinvertebrados, y que estos no son susceptibles de muestreo con la red, es necesario complementar con la extracción manual de rocas y piezas grandes de madera muerta. Al igual que el muestreo con pateo solo se colecta hasta la repetición de taxones, con la intención de evitar el sobremuestreo y la sobre extracción de organismos. En general la colecta incluye unos 100 individuos aproximadamente; con este número se supone que ya no aparecen nuevos taxones; por lo tanto, es importante vigilar durante la selección de los individuos no colectar más allá de este número si no es necesario, en términos de representatividad numérica y de la diversidad del sitio. </p>
          <p align="justify"> Los organismos deben ser almacenados en frascos de plástico previamente etiquetados con el nombre de la localidad y la fecha, con alcohol al 70%; esto con la finalidad de preservar todos los organismos y evitar pérdidas en el transporte por la presencia de depredadores. </p>          
          <h1 class="display-6">Análisis</h1>
          <img src="../assets/muestra_bemeqi.png" class="img-fluid  hover-shadow" alt="diatomea" style="float:right; height: 166px; margin-left: 20px" />
          <p align="justify" > Una vez en el laboratorio, las muestras deben ser revisadas al microscopio estereoscópico para ser clasificadas taxonómicamente en las diferentes familias.</p>
          <p align="justify" > Es recomendable vaciar la muestra que viene almacenada en el frasco en una caja de Petri, e incluso es posible ir separando visualmente todos los organismos que en un inicio se ven idénticos. Esta separación inicial facilita la revisión al microscopio y la identificación de las familias para los conteos de abundancias absolutas. </p>
          <h1 class="display-6">Evaluación</h1>
          <p align="justify"> Respecto a los ensambles que son característicos de los distintos estados de calidad ecológica, los macroinvertebrados nos permiten detectar 3 tipos de sitios: </p> 
          <ol style = "text-align: justify">
            <li align="justify">Sitios en muy buen estado de calidad físicoquímica y del hábitat.</li>
            <li align="justify">Sitios con calidad moderada, ya que la calidad físicoquímica o la calidad del hábitat presentan alguna afectación. Estos sitios son intermedios y podrían moverse con facilidad hacia los de mala calidad si las perturbaciones son persistentes en el tiempo. </li>
            <li align="justify">Los sitios con mala calidad, en particular son sitios alterados con altas concentraciones de nutrientes y afectaciones como canalizaciones e influencia urbana cercana. </li>
          </ol>
          <p align="justify"> A continuación se presentan los valores de  especificidad (<i>v</i> ) y sensibilidad(<i>s</i>) por taxa:</p> 
          <MDBTable borderless responsive>
            <thead class="table-info">
              <tr>
                <th scope="col">Taxa</th>
                <th scope="col">Especificidad (<i>v </i>)</th>
                <th scope="col">Sensibilidad (<i>s </i>)</th>
              </tr>
            </thead>
            <tbody>
              <tr  class="table-active">
                <th scope="row"> <b>Clase Arachnida</b> &emsp; &emsp; &emsp; &emsp; </th> <td></td> <td></td> 
              </tr>
              <tr  class="table-active">
                <th scope="row">Orden Trombidiformes &emsp; </th> <td></td> <td></td> 
              </tr>
              <tr>
                <td scope="row"><i>Hydracarina </i> </td>
                <td>2</td>
                <td>4</td>
              </tr>
              <tr  class="table-active">
                <th scope="row"> <b>Clase Clitellata </b> &emsp; &emsp; &emsp; &emsp; </th> <td></td> <td></td> 
              </tr>
              <tr >
                <td scope="row"><i>Subclase Oligochaeta  </i> </td>
                <td>3</td>
                <td>4</td>
              </tr>
              <tr  class="table-active">
                <th scope="row"> <b>Clase Insecta  </b> &emsp; &emsp; &emsp; &emsp; </th> <td></td> <td></td> 
              </tr>
              <tr  class="table-active">
                <th scope="row">Orden Coleoptera  &emsp; </th> <td></td> <td></td> 
              </tr>
              <tr>
                <td scope="row"><i>F. Dytiscidae </i> </td>
                <td>3</td>
                <td>2.5</td>
              </tr>
              <tr>
                <td scope="row"><i>F. Elmidae </i> </td>
                <td>2</td>
                <td>4</td>
              </tr>
              <tr  class="table-active">
                <th scope="row">Orden Diptera  &emsp; </th> <td></td> <td></td> 
              </tr>
              <tr>
                <td scope="row"><i>F. Ceratopogonidae  </i> </td>
                <td>3</td>
                <td>1</td>
              </tr>
              <tr>
                <td scope="row"><i>F. Chironomidae </i></td>
                <td>3</td>
                <td>4</td>
              </tr>
              <tr>
                <td scope="row"><i>F. Dixidae  </i> </td>
                <td>2</td>
                <td>2.5</td>
              </tr>
              <tr>
                <td scope="row"><i>F. Ephydridae </i> </td>
                <td>3</td>
                <td>4</td>
              </tr>
              <tr>
                <td scope="row"><i>F. Limoniidae </i></td>
                <td>2</td>
                <td>2.5</td>
              </tr>
              <tr>
                <td scope="row"><i>F. Simuliidae </i> </td>
                <td>2</td>
                <td>2.5</td>
              </tr>
              <tr>
                <td scope="row"><i>F. Tipulidae </i> </td>
                <td>3</td>
                <td>2.5</td>
              </tr>
              <tr  class="table-active">
                <th scope="row">Orden Ephemeroptera   &emsp; </th> <td></td> <td></td> 
              </tr>
              <tr>
                <td scope="row"><i>F. Ameletidae  </i> </td>
                <td>1</td>
                <td>2.5</td>
              </tr>
              <tr>
                <td scope="row"><i>F. Baetidae  </i> </td>
                <td>3</td>
                <td>4</td>
              </tr>
              <tr>
                <td scope="row"><i>F. Heptageniidae  </i> </td>
                <td>2</td>
                <td>1</td>
              </tr>
              <tr  class="table-active">
                <th scope="row">Orden Plecoptera &emsp; </th>  <td></td>  <td></td>
              </tr>
              <tr>
                <td scope="row"><i>F. Nemouridae  </i> </td>
                <td>2</td>
                <td>1</td>
              </tr>
              <tr>
                <td scope="row"><i>F. Perlodidae   </i> </td>
                <td>1</td>
                <td>1</td>
              </tr>
              <tr  class="table-active">
                <th scope="row">Orden Trichoptera &emsp; </th>    <td></td><td></td> 
              </tr>
              <tr>
                <td scope="row"><i>F. Glossosomatidae   </i> </td>
                <td>2</td>
                <td>1</td>
              </tr>
              <tr>
                <td scope="row"><i>F. Helicopsychidae   </i> </td>
                <td>1</td>
                <td>1</td>
              </tr>
              <tr>
                <td scope="row"><i>F. Hydrobiosidae   </i> </td>
                <td>2</td>
                <td>2.5</td>
              </tr>
              <tr>
                <td scope="row"><i>F. Hydrophilidae   </i> </td>
                <td>2</td>
                <td>4</td>
              </tr>
              <tr>
                <td scope="row"><i>F. Hydroptilidae   </i> </td>
                <td>2</td>
                <td>2.5</td>
              </tr>
              <tr>
                <td scope="row"><i>F. Hydropsychidae   </i> </td>
                <td>2</td>
                <td>2.5</td>
              </tr>
              <tr>
                <td scope="row"><i>F. Lepidostomatidae   </i> </td>
                <td>1</td>
                <td>1</td>
              </tr>
              <tr>
                <td scope="row"><i>F. Leptoceridae    </i> </td>
                <td>2</td>
                <td>2.5</td>
              </tr>
              <tr>
                <td scope="row"><i>F. Limnephilidae   </i> </td>
                <td>3</td>
                <td>2.5</td>
              </tr>
              <tr>
                <td scope="row"><i>F. Polycentropodidae    </i> </td>
                <td>2</td>
                <td>2.5</td>
              </tr>
              <tr>
                <td scope="row"><i>F. Xiphocentronidae    </i> </td>
                <td>2</td>
                <td>2.5</td>
              </tr>
              <tr  class="table-active">
                <th scope="row"> <b>Clase Malacostraca   </b> &emsp; &emsp; &emsp; &emsp; </th> <td></td> <td></td> 
              </tr>
              <tr >
                <th scope="row">Orden Amphipoda   &emsp; </th> <td>2</td> <td>4</td> 
              </tr>
              <tr  class="table-active">
                <th scope="row"> <b>Clase Rhabditophora    </b> &emsp; &emsp; &emsp; &emsp; </th> <td></td> <td></td> 
              </tr>
              <tr  class="table-active">
                <th scope="row">Orden Amphipoda   &emsp; </th> <td></td> <td></td> 
              </tr>
              <tr>
                <td scope="row"><i>F. Duguesiidae     </i> </td>
                <td>3</td>
                <td>4</td>
              </tr>
            </tbody>
          </MDBTable>
          <p align="justify" > El cálculo del valor de sensibilidad y especificidad propuestos para cada una de las familias de macroinvertebrados son usados en la formulación matemática propuesta en <a href="https://www.jlimnol.it/index.php/jlimnol/article/view/jlimnol.2015.1304" target="_blank" rel="noopener noreferrer"> Caro <i>et al.</i> ( 2015 ) </a> para calcular el BEMEQI : </p>           
          <vue-mathjax :formula="formula" />  <br>
          <p align="justify">El valor del BEMEQI se clasifica en tres categorías definidas como: alta, moderada, y mala, donde los valores inferiores a 1.66 pueden considerarse Sitios con todos los elementos de calidad hidromorfológica en condiciones óptimas y sin incremento de nutrientes. Los valores entre 1.67 y 3.5 indican sitios que presentan algún tipo de alteración, bien sea relacionada con la calidad hidromorfológica o con el enriquecimiento orgánico. Suelen ser sitios con un incremento notable en el fósforo y/o nitrógeno inorgánico disuelto. Los sitios con BEMEQI mayor a 3.5 tienen alterada la calidad hidromorfológica usualmente por la presencia de presas, derivaciones, canalizaciones y por la pérdida de la vegetación ribereña la cual es reemplazada por cultivos o zonas urbanas muy cercanas.</p>
          <MDBTable border="gray" bordered>
            <thead >
              <tr>
                <th scope="col">Calidad ecológica</th>
                <th scope="col">BEMEQI</th>
              </tr>
            </thead>
            <tbody>
              <tr class="table-primary">
                <td > Alta </td>
                <td> 1 - 1.66</td>
              </tr>
              <tr class="table-warning">
                <td > Moderada </td>
                <td> 1.67 - 3.5</td>
              </tr>
              <tr class="table-danger">
                <td > Mala </td>
                <td> > 3.5 </td>
              </tr>
            </tbody>
          </MDBTable>
        </MDBTabPane>

        <!-- Calculo con microservicio-->
        <MDBTabPane tabId="calculo-bemeqi">
        <h1 class="display-5">Cálculo del protocolo</h1>
        <br> <p align="justify"> Aquí puedes descargar el formato para el cálculo del índice y llenarlo con datos del conteo de macroalgas obtenidos en la colecta, debajo del botón encontrarás información de cómo llenar adecuadamente el formato. </p>
        <div class="d-grid gap-2">
          <MDBBtn color="primary" download @click="genera">
            Generar 
          </MDBBtn> 
        </div>
        <!-- Imagen de ejemplo de llenado -->
          <br><h1 class="display-6">Hoja 1: BEMEQI</h1>
          <img src="../assets/hoja1BEMEQI.png"  class="img-fluid  hover-shadow" alt="hoja1" style=" height: 252px; " /> <br> <br>
          <p align="justify">En esta hoja, hasta arriba, deberás anotar los datos del proyecto para el que se realiza este cálculo, los nombres de los autores y revisor, y las fechas de colecta y revisión. </p>
          <p align="justify">Posteriormente según los datos analizados de la colecta deberás anotar el conteo de cada taxón reconocido(valor <i>a</i>), según su acomodo por filas.</p>
          <p align="justify">Cada columna, a partir de la columna E, representa un sitio de colecta involucrado en el proyecto. De inicio se incluyen 20 sitios numerados pero se puede cambiar el nombre de las columnas por el nombre del sitio recolectada, o bien, asignar claves a cada sitio y ponerlas en dichas columnas. </p>
          <p align="justify">Al momento de escribir los datos de conteo de especies la hoja hará los cálculos actualizados al momento, desplegando el valor BEMEQI de cada sitio en la parte de abajo de la hoja en las filas 39-41 con los códigos de colores descritos en la pestaña INFORMACIÓN DEL PROTOCOLO en este sitio; y hasta arriba de la hoja se mostrará también el valor BEMEQI promedio de todos los sitios que tengan datos introducidos, así como informar la calidad ecológica en la que se pueden clasificar los sitos introducidos.  </p>          
        </MDBTabPane>
      </MDBTabContent>
      <!-- Tabs content -->
    </MDBTabs>
  </div>
</div>

</template>

<style scoped>
  * {
    box-sizing: border-box;
  }

  .row {
    padding-top: 50px;
    display: flex;
  }

  /* Create two columns that sits next to each other */
  .column {
    float: left;
    width: 25%;
    padding: 10px;
    padding-top: 25px;
  }
  .column-2 {
    float: left;
    width: 75%;
    padding: 40px;
    }

  /* Clear floats after the columns */
  .row:after {
    content: "";
    display: table;
    clear: both;
  }
  /* Responsive layout - makes the two columns stack on top of each other instead of next to each other */
  @media screen and (max-width: 600px) {
    .column {
      width: 100%;
    }
    .column-2 {
      width: 100%;
    }
  }

  #inner {
    display: table;
    margin: 0 auto;
  }

</style>

<script >
  import { MDBListGroup, 
           MDBListGroupItem,
           MDBTabs,
           MDBTabNav,
           MDBTabContent,
           MDBTabItem,
           MDBTabPane,
           MDBTable,
           MDBBtn,
          } from "mdb-vue-ui-kit";
  import { ref } from 'vue';
  import router from '../router'
  export default {
    components: {
      MDBListGroup,
      MDBListGroupItem,
      MDBTabs,
      MDBTabNav,
      MDBTabContent,
      MDBTabItem,
      MDBTabPane,
      MDBTable,
      MDBBtn,
    },
    setup() {
      const activeTabId = ref('info-bemeqi');
     
      return {
        activeTabId,
        formula: '$$BEMEQI = \\frac{ \\sum_{j=1}^n a_j s_j v_j } { \\sum_{j=1}^n a_j v_j} $$',
      };
    },
    methods: {
      navega: function(url) {
        router.push(url).catch(()=>{});
        },
      genera(){
        console.log("se llega aqui")
        const link = document.createElement('a');
        link.href = './assets/files/ProtocoloBEMEQI.xlsx';
        link.setAttribute('download', 'ProtocoloBEMEQI.xlsx'); 
        document.body.appendChild(link);
        link.click();
      }
    } 
  };
</script>
