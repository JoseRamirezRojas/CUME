<template>

<div class="row">
  <div class="column" style="background-color: #e3f2fd">
    <MDBListGroup  id="inner" >
      <MDBListGroupItem tag="button" @click="navega('/protocolos-info')" action >Protocolos</MDBListGroupItem>
      <MDBListGroupItem tag="button" @click="navega('/protocolo-fisicoquimica')" action> Calidad fisicoquímica </MDBListGroupItem>
      <MDBListGroupItem tag="button" action active> HYQI </MDBListGroupItem>
      <MDBListGroupItem tag="button" @click="navega('/protocolo-deqi')" action>DEQI</MDBListGroupItem>
      <MDBListGroupItem tag="button" @click="navega('/protocolo-meqi')" action>MEQI</MDBListGroupItem>
      <MDBListGroupItem tag="button" @click="navega('/cuenca-conservacion')" action>BEMEQI </MDBListGroupItem>
      <MDBListGroupItem tag="button" @click="navega('/cuenca-jurisdiccion')" action >Material de trabajo</MDBListGroupItem>
    </MDBListGroup>

  </div>

  <div class="column-2" style="background-color:#F8F8F8;">
    <MDBTabs v-model="activeTabId">
      <!-- Tabs navs -->
      <MDBTabNav pills justify tabsClasses="mb-3">
        <MDBTabItem tabId="info-hyqi" href="info-hyqi"      >Información del protocolo</MDBTabItem>
        <MDBTabItem tabId="calculo-hyqi" href="calculo-hyqi">Cálculo de índice HYQI</MDBTabItem>
      </MDBTabNav>
      <!-- Tabs navs -->
      <!-- Tabs content -->
      <MDBTabContent>
        <!-- info protocolo-->
        <MDBTabPane tabId="info-hyqi">
          <p align="justify" > El artículo 3, fracción XLVII, de la <a href="https://www.diputados.gob.mx/LeyesBiblio/pdf/16_060120.pdf" target="_blank" rel="noopener noreferrer" > Ley de Aguas Nacionales </a>, define como “Ribera o Zona Federal” a las fajas de diez metros de anchura contiguas al cauce de las corrientes o al vaso de los depósitos de propiedad nacional, medidas horizontalmente a partir del nivel de aguas máximas ordinarias. </p>
          <p align="justify">El ecosistema de ribera es un elemento importante en el funcionamiento de un río, sus dimensiones laterales y afluentes, así como muchas de sus dimensiones verticales, están contenidas en su hábitat. Provee servicios a las poblaciones como el suministro de agua, para consumo humano y para el desarrollo de agricultura, ganadería, industria; purificación y transporte de agua, sedimentos y nutrientes; control de inundaciones y sequías; hábitat para las especies; retención de suelo y valores escénicos, culturales y de recreación </p>  
          <h1 class="display-6">Calidad Hidromorfológica</h1>
          <p align = "justify"> El protocolo HYQI (por sus siglas en inglés HYdromorphological Quality Index) incluye datos generales sobre el sitio de muestreo; registro de parámetros físicos y químicos del agua; porcentaje de la composición del sustrato; tipo de bosque del sitio de muestreo, forma de vida predominante. El protocolo contiene doce parámetros hidromorfológicos a evaluar, divididos en tres secciones (las tablas provistas ayudan al reconocimiento de cada uno de los parámetros descritos). </p>
          <p align="justify"  class="fs-4">Sección 1: La calidad de la cuenca</p>
          <p align="justify" >  Se examina la cobertura vegetal, la estabilidad del banco, talla del sustrato y desarrollo de ganadería y agricultura en la zona de ribera:</p>
          <MDBTable borderless responsive>
            <thead class="table-info">
              <tr>
                <th scope="col">Banco estable </th>
                <th scope="col">Banco parcialmente erosionado</th>
                <th scope="col">Banco totalmente erosionado</th>
              </tr>
            </thead>
            <tbody>
              <tr>
                <td > <img src="../assets/banco_estable_hyqi.jpg" alt="banco estable" 
                  style="width: 163px; height: 126px;" class="rounded" /> </td>
                <td><img src="../assets/banco_parc_estable.jpg" alt="banco parcialmente estable" 
                  style="width: 160px; height: 126px;" class="rounded" /></td>
                <td><img src="../assets/banco_tot_erosionado.jpg" alt="banco erosionado" 
                  style="width: 160px; height: 126px;" class="rounded" /></td>
              </tr>
            </tbody>
          </MDBTable>
          <p align="justify" > Las características del sustrato relaciona la heterogeneidad del sustrato orgánico e inorgánico y define el número de microambientes potenciales en los que se pueden desarrollar organismos acuáticos:</p>
          <MDBTable borderless responsive>
            <thead class="table-info">
              <tr>
                <th scope="col">Sustratos inorgánicos (arena, grava, roca emergente) </th>
                <th scope="col">Sustratos orgánicos (algas, vegetación acuática)</th>
              </tr>
            </thead>
            <tbody>
              <tr>
                <td > <img src="../assets/sust_inorg.jpg" alt="inorganico" 
                  style="width: 196px; height: 129px;" class="rounded" /> </td>
                <td><img src="../assets/sust_organ.jpg" alt="organico" 
                  style="width: 196px; height: 129px;" class="rounded" /></td>
              </tr>
            </tbody>
          </MDBTable>
          <p align="justify"> Las principales intervenciones humanas en los ríos son por la presencia de trucheros, actividades agrícolas y ganaderas:</p>
          <MDBTable borderless responsive>
            <thead class="table-info">
              <tr>
                <th scope="col">Sin presencia de cultivos o zonas para el ganado </th>
                <th scope="col">+80% de suelo para uso piscícola, agrícola y ganadero</th>
              </tr>
            </thead>
            <tbody>
              <tr>
                <td > <img src="../assets/sin_ganaderia.jpg" alt="buen suelo" 
                  style="width: 196px; height: 129px;" class="rounded" /> </td>
                <td><img src="../assets/con_ganaderia.jpg" alt="suelo usado" 
                  style="width: 196px; height: 129px;" class="rounded" /></td>
              </tr>
            </tbody>
          </MDBTable>
          <p align="justify" class="fs-4">Sección 2: La hidrología del río</p>
          <p align="justify" >  Se examina la hidrología según la naturalidad del canal, se tienen en cuenta las interrupciones del flujo por la construcción de presas.</p>
          <MDBTable borderless responsive>
            <thead class="table-info">
              <tr>
                <th scope="col">Sin presencia de  presas</th>
                <th scope="col">Presencia de presas</th>
              </tr>
            </thead>
            <tbody>
              <tr>
                <td > <img src="../assets/sin_presas.jpg" alt="sin presa" 
                  style="width: 196px; height: 129px;" class="rounded" /> </td>
                <td><img src="../assets/con_presas.jpg" alt="con presa" 
                  style="width: 196px; height: 129px;" class="rounded" /></td>
              </tr>
            </tbody>
          </MDBTable>
          <p align="justify" >El régimen de velocidad y profundidad se determina observando la presencia de cuatro combinaciones posibles en el lecho del río: rápido-somero, rápido-profundo, lento-somero y lento-profundo. Una zona somera es aquella con profundidades menores a 0.4 <i>m</i> y una zona rápida está por encima de los 0.3 <i>m/s</i> </p>
          <img src="../assets/reg_velprof.jpg" height="220" width="570" class="img-fluid  hover-shadow" alt="ejemplo aforo" />  <br><br>
          <p align="justify"> La alteración del canal se puede reconocer por la construcción de infraestructura aledaña al margen de ambos lados del río.</p>
          <MDBTable borderless responsive>
            <thead class="table-info">
              <tr>
                <th scope="col">Sin alteraciones en el canal</th>
                <th scope="col">Alteraciones parciales</th>
                <th scope="col">Canal totalmente modificado</th>
              </tr>
            </thead>
            <tbody>
              <tr>
                <td > <img src="../assets/sin_alteracion.jpg" alt="nada alterado" 
                  style="width: 163px; height: 126px;" class="rounded" /> </td>
                <td><img src="../assets/alteracion_parc.jpg" alt="Alteraciones parciales" 
                  style="width: 160px; height: 126px;" class="rounded" /></td>
                  <td><img src="../assets/alteracion_total.jpg" alt="todo alterado" 
                  style="width: 160px; height: 126px;" class="rounded" /></td>
              </tr>
            </tbody>
          </MDBTable>
          <p align="justify"> El estado del canal se reconoce por la cantidad del agua que circula y el nivel que ocupa en ambos bancos de la orilla. </p>
          <MDBTable borderless responsive>
            <thead class="table-info">
              <tr>
                <th scope="col">El agua llega a la base de los bancos, sustrato no expuesto</th>
                <th scope="col">Muy poca agua en el canal</th>
              </tr>
            </thead>
            <tbody>
              <tr>
                <td > <img src="../assets/buen_estado.jpg" alt="sin presa" 
                  style="width: 196px; height: 129px;" class="rounded" /> </td>
                <td><img src="../assets/mal_estado.jpg" alt="con presa" 
                  style="width: 196px; height: 129px;" class="rounded" /></td>
              </tr>
            </tbody>
          </MDBTable>
          <p align="justify" class="fs-4">Sección 3: Las perturbaciones antropogénicas</p>
          <p align="justify" > Estas tienen que ver con la contaminación del agua por efluentes urbanos, desarrollo humano y contaminación.</p>
          <MDBTable borderless responsive>
            <thead class="table-info">
              <tr>
                <th scope="col">Sin efluentes al río</th>
                <th scope="col">Descarga directa de residuos al río</th>
              </tr>
            </thead>
            <tbody>
              <tr>
                <td > <img src="../assets/sin_efluentes.jpg" alt="sin efluentes" 
                  style="width: 196px; height: 129px;" class="rounded" /> </td>
                <td><img src="../assets/descarga_eflu.jpg" alt="descargas al rio" 
                  style="width: 196px; height: 129px;" class="rounded" /></td>
              </tr>
            </tbody>
          </MDBTable>
          <p align="justify">El desarrollo de infraestructura urbana refiere a la presencia de caminos, carreteras o puentes asociadas al canal del río, incluyendo extracción en este.</p>
          <MDBTable borderless responsive>
            <thead class="table-info">
              <tr>
                <th scope="col">Sin infraestructura</th>
                <th scope="col">Carreteras y construcciones</th>
              </tr>
            </thead>
            <tbody>
              <tr>
                <td > <img src="../assets/sin_construcc.jpg" alt="sin construiir" 
                  style="width: 196px; height: 129px;" class="rounded" /> </td>
                <td><img src="../assets/construcciones.jpg" alt="construcciones" 
                  style="width: 196px; height: 129px;" class="rounded" /></td>
              </tr>
            </tbody>
          </MDBTable>
          <MDBTable borderless responsive>
            <thead class="table-info">
              <tr>
                <th scope="col">Sin extracción</th>
                <th scope="col">Extracción directa en la ribera</th>
              </tr>
            </thead>
            <tbody>
              <tr>
                <td > <img src="../assets/no_extracc.jpg" alt="sin extracción" 
                  style="width: 196px; height: 129px;" class="rounded" /> </td>
                <td><img src="../assets/extracc_directa.jpg" alt="extracción directa" 
                  style="width: 196px; height: 129px;" class="rounded" /></td>
              </tr>
            </tbody>
          </MDBTable>
          <p align="justify">Se examina también la contaminación de residuos sólidos por basura y escombros que son vertidos al lecho del río.</p>
          <MDBTable borderless responsive>
            <thead class="table-info">
              <tr>
                <th scope="col">Sin basura, y sin escombros</th>
                <th scope="col">Presencia difusa de basura y escombros</th>
                <th scope="col">Presencia abundante de basura y escombros</th>
              </tr>
            </thead>
            <tbody>
              <tr>
                <td > <img src="../assets/sin_basura.jpg" alt="nada basura" 
                  style="width: 163px; height: 135px;" class="rounded" /> </td>
                <td><img src="../assets/poca_basura.jpg" alt="poca basura" 
                  style="width: 160px; height: 135px;" class="rounded" /></td>
                  <td><img src="../assets/mucha_basura.jpg" alt="basura y escombros" 
                  style="width: 160px; height: 135px;" class="rounded" /></td>
              </tr>
            </tbody>
          </MDBTable>
          <br> <br>
          <h1 class="display-6">Evaluación</h1> <br>
          <p align="justify" >El protocolo HYQI tiene un total de 120 puntos y cada sección tiene un total de 40 puntos. El puntaje parcial permitirá identificar la sección mejor conservada de las 3 descritas anteriormente. 
          El puntaje final es el resultado de la suma de los doce parámetros analizados y se verifica su estado de acuerdo al puntaje final y una tabla de colores. Si el valor supera los 85 puntos está en condición de ser considerado como sitio de referencia por su buena calidad hidromorfológica. </p>
          <MDBTable border="dark">
            <thead >
              <tr>
                <th scope="col">Calidad hidromorfológica</th>
                <th scope="col">Puntaje</th>
              </tr>
            </thead>
            <tbody>
              <tr class="table-primary">
                <td > Óptima </td>
                <td> 120-85</td>
              </tr>
              <tr class="table-success">
                <td > Media </td>
                <td> 84-47</td>
              </tr>
              <tr class="table-warning">
                <td > Mala </td>
                <td> 46-13</td>
              </tr>
              <tr class="table-danger">
                <td > Pobre </td>
                <td> &lt;12 </td>
              </tr>
            </tbody>
          </MDBTable>

        </MDBTabPane>

        <!-- Calculo con microservicio-->
        <MDBTabPane tabId="calculo-hyqi">
        <h1 class="display-5">Cálculo del protocolo</h1>
        <br> <p align="justify"> Aquí puedes introducir los parámetros hidromorfológicos del agua obtenidos en la colecta, hasta abajo encontrarás un ejemplo del llenado de esta ficha. Al seleccionar una opción en cada parámetro se despliega más información sobre esta. Al dar clic en Generar se descargará un archivo editable de Excel donde puedes visualizar los datos y el índice HYQI calculado, así como una guía en la hoja 2 del archivo para  mayor referencia en caso de querer hacer cambios en la hoja de cálculo obtenida.</p>
        <p align="justify"><b> Se recomienda rotar dispositivos móviles en posición horizontal</b> </p>
        <MDBRow tag="form" class="needs-validation " @submit.prevent="genera" >
          <MDBRow> 
            <MDBCol >
              <h5> <b>Cuenca</b> </h5>
            </MDBCol>
          </MDBRow>
          <MDBRow> 
            <MDBCol >
              <h6>Cobertura vegetal <br/> (% de suelo)</h6>
            </MDBCol>   <MDBCol ></MDBCol>
          </MDBRow>
          <MDBRow> 
            <MDBCol >
              <p>Ribera derecha</p>
            </MDBCol>
            <MDBCol >
              <MDBBtnGroup class="shadow-0">
                <MDBCol >
                  <MDBTooltip v-model="tooltip1" direction="bottom">
                    <template #reference>
                      <MDBRadio required v-model="form.coberturaDer"  :btnCheck="true" :wrap="false" labelClass="btn btn-info"
                        label=">70% " name="coberturaDer" value="5" />
                    </template>
                    <template #tip> Óptimo</template>
                  </MDBTooltip>
                </MDBCol>
                <MDBCol>
                  <MDBTooltip v-model="tooltip2" direction="bottom">
                    <template #reference>
                      <MDBRadio required v-model="form.coberturaDer" :btnCheck="true" :wrap="false" labelClass="btn btn-info"
                      label=" 70-50%" name="coberturaDer" value="3"  />
                    </template>
                    <template #tip> Medio</template>
                  </MDBTooltip>
                </MDBCol>
                <MDBCol >
                  <MDBTooltip v-model="tooltip3" direction="bottom">
                    <template #reference>
                      <MDBRadio required v-model="form.coberturaDer" :btnCheck="true" :wrap="false" labelClass="btn btn-info"
                      label=" 50-30%" name="coberturaDer" value="2" />
                      </template>
                    <template #tip> Malo</template>
                  </MDBTooltip>
                </MDBCol>
                <MDBCol >
                  <MDBTooltip v-model="tooltip4" direction="bottom">
                    <template #reference>
                      <MDBRadio required v-model="form.coberturaDer" :btnCheck="true" :wrap="false" labelClass="btn btn-info"
                      label="<30%" name="coberturaDer" value="0" />
                    </template>
                    <template #tip> Pobre</template>
                  </MDBTooltip>
                </MDBCol>
              </MDBBtnGroup>
            </MDBCol>
          </MDBRow>
          <MDBRow> 
            <MDBCol >
              <p>Ribera izquierda</p>
            </MDBCol>
            <MDBCol >
              <MDBBtnGroup class="shadow-0">
                <MDBCol >
                  <MDBTooltip v-model="tooltip5" direction="bottom">
                    <template #reference>
                      <MDBRadio required v-model="form.coberturaIzq"  :btnCheck="true" :wrap="false" labelClass="btn btn-info"
                        label=">70% " name="coberturaIzq" value="5" />
                    </template>
                    <template #tip> Óptimo</template>
                  </MDBTooltip>
                </MDBCol>
                <MDBCol>
                  <MDBTooltip v-model="tooltip6" direction="bottom">
                    <template #reference>
                      <MDBRadio required v-model="form.coberturaIzq" :btnCheck="true" :wrap="false" labelClass="btn btn-info"
                      label=" 70-50%" name="coberturaIzq" value="3"  />
                    </template>
                    <template #tip> Medio</template>
                  </MDBTooltip>
                </MDBCol>
                <MDBCol >
                  <MDBTooltip v-model="tooltip7" direction="bottom">
                    <template #reference>
                      <MDBRadio required v-model="form.coberturaIzq" :btnCheck="true" :wrap="false" labelClass="btn btn-info"
                      label=" 50-30%" name="coberturaIzq" value="2" />
                      </template>
                    <template #tip> Malo</template>
                  </MDBTooltip>
                </MDBCol>
                <MDBCol >
                  <MDBTooltip v-model="tooltip8" direction="bottom">
                    <template #reference>
                      <MDBRadio required v-model="form.coberturaIzq" :btnCheck="true" :wrap="false" labelClass="btn btn-info"
                      label="<30%" name="coberturaIzq" value="0" />
                    </template>
                    <template #tip> Pobre</template>
                  </MDBTooltip>
                </MDBCol>
              </MDBBtnGroup>
            </MDBCol>
          </MDBRow>
          <MDBRow> 
            <MDBCol > <h6>Estabilidad del banco  <br/> (% de erosión)</h6> </MDBCol>
            <MDBCol >
              <MDBBtnGroup class="shadow-0">
                <MDBCol >
                  <MDBTooltip v-model="tooltip9" direction="bottom">
                    <template #reference>
                      <MDBRadio required v-model="form.estabilidad"  :btnCheck="true" :wrap="false" labelClass="btn btn-info"
                        label="<10% " name="estabilidad" value="10" />
                    </template>
                    <template #tip> Estable</template>
                  </MDBTooltip>
                </MDBCol>
                <MDBCol>
                  <MDBTooltip v-model="tooltip10" direction="bottom">
                    <template #reference>
                      <MDBRadio required v-model="form.estabilidad" :btnCheck="true" :wrap="false" labelClass="btn btn-info"
                      label=" 10-50%" name="estabilidad" value="7"  />
                    </template>
                    <template #tip> Moderadamente estable</template>
                  </MDBTooltip>
                </MDBCol>
                <MDBCol >
                  <MDBTooltip v-model="tooltip11" direction="bottom">
                    <template #reference>
                      <MDBRadio required v-model="form.estabilidad" :btnCheck="true" :wrap="false" labelClass="btn btn-info"
                      label=" 50-80%" name="estabilidad" value="4" />
                      </template>
                    <template #tip> Moderadamente inestable</template>
                  </MDBTooltip>
                </MDBCol>
                <MDBCol >
                  <MDBTooltip v-model="tooltip12" direction="bottom">
                    <template #reference>
                      <MDBRadio required v-model="form.estabilidad" :btnCheck="true" :wrap="false" labelClass="btn btn-info"
                      label=">80%" name="estabilidad" value="1" />
                    </template>
                    <template #tip> Inestable</template>
                  </MDBTooltip>
                </MDBCol>
              </MDBBtnGroup>
            </MDBCol>
          </MDBRow>
          <MDBRow class="overflow-hidden"> 
            <MDBCol >
              <h6>Característ. del sustrato <br/> (materiales, vegetación acuática)</h6>
            </MDBCol>
            <MDBCol >
              <MDBBtnGroup class="shadow-0">
                <MDBCol >
                  <MDBTooltip v-model="tooltip13" direction="bottom">
                    <template #reference>
                      <MDBRadio required v-model="form.sustrato"  :btnCheck="true" :wrap="false" labelClass="btn btn-info"
                        label="Óptimo" name="sustrato" value="10" />
                    </template>
                    <template #tip> Grava y arena; raíces sumergidas y veget. acuática</template>
                  </MDBTooltip>
                </MDBCol>
                <MDBCol>
                  <MDBTooltip v-model="tooltip14" direction="bottom">
                    <template #reference>
                      <MDBRadio required v-model="form.sustrato" :btnCheck="true" :wrap="false" labelClass="btn btn-info"
                      label=" Medio" name="sustrato" value="7"  />
                    </template>
                    <template #tip>Arena, arcilla o lodo; algunas raíces sumerg. y veget. acuática</template>
                  </MDBTooltip>
                </MDBCol>
                <MDBCol >
                  <MDBTooltip v-model="tooltip15" direction="bottom">
                    <template #reference>
                      <MDBRadio required v-model="form.sustrato" :btnCheck="true" :wrap="false" labelClass="btn btn-info"
                      label="Malo" name="sustrato" value="4" />
                      </template>
                    <template #tip> Arcilla en superficie, pocas raíces sumerg. y sin vegetación</template>
                  </MDBTooltip>
                </MDBCol>
                <MDBCol >
                  <MDBTooltip v-model="tooltip16" direction="bottom">
                    <template #reference>
                      <MDBRadio required v-model="form.sustrato" :btnCheck="true" :wrap="false" labelClass="btn btn-info"
                      label="Pobre" name="sustrato" value="1" />
                    </template>
                    <template #tip> Capa de arcilla o rocas, sin raíces ni vegetación </template>
                  </MDBTooltip>
                </MDBCol>
              </MDBBtnGroup>
            </MDBCol>
          </MDBRow>
          <MDBRow> 
            <MDBCol>
              <h6>Desarrollo de ganadería y agricultura <br/> (% uso de suelo)</h6>
            </MDBCol>   <MDBCol ></MDBCol>
          </MDBRow>
          <MDBRow> 
            <MDBCol >
              <p>Ribera derecha</p>
            </MDBCol>
            <MDBCol >
              <MDBBtnGroup class="shadow-0">
                <MDBCol >
                  <MDBTooltip v-model="tooltip17" direction="bottom">
                    <template #reference>
                      <MDBRadio required v-model="form.agriculturaDer"  :btnCheck="true" :wrap="false" labelClass="btn btn-info"
                        label="0%" name="agriculturaDer" value="5" />
                    </template>
                    <template #tip> Sin cultivos o zonas para ganado ni derivaciones para estas actividades</template>
                  </MDBTooltip>
                </MDBCol>
                <MDBCol>
                  <MDBTooltip v-model="tooltip18" direction="bottom">
                    <template #reference>
                      <MDBRadio required v-model="form.agriculturaDer" :btnCheck="true" :wrap="false" labelClass="btn btn-info"
                      label=" 20%" name="agriculturaDer" value="3"  />
                    </template>
                    <template #tip> Suelo para uso agrícola y ganadero</template>
                  </MDBTooltip>
                </MDBCol>
                <MDBCol >
                  <MDBTooltip v-model="tooltip19" direction="bottom">
                    <template #reference>
                      <MDBRadio required v-model="form.agriculturaDer" :btnCheck="true" :wrap="false" labelClass="btn btn-info"
                      label=" 50%" name="agriculturaDer" value="2" />
                      </template>
                    <template #tip> Suelo para uso agrícola y ganadero</template>
                  </MDBTooltip>
                </MDBCol>
                <MDBCol >
                  <MDBTooltip v-model="tooltip20" direction="bottom">
                    <template #reference>
                      <MDBRadio required v-model="form.agriculturaDer" :btnCheck="true" :wrap="false" labelClass="btn btn-info"
                      label=">80%" name="agriculturaDer" value="0" />
                    </template>
                    <template #tip> Uso agrícola y ganadero, con derivaciones para uso doméstico o industrial</template>
                  </MDBTooltip>
                </MDBCol>
              </MDBBtnGroup>
            </MDBCol>
          </MDBRow>
          <MDBRow> 
            <MDBCol >
              <p>Ribera izquierda</p>
            </MDBCol>
            <MDBCol >
              <MDBBtnGroup class="shadow-0">
                <MDBCol >
                  <MDBTooltip v-model="tooltip21" direction="bottom">
                    <template #reference>
                      <MDBRadio required v-model="form.agriculturaIzq"  :btnCheck="true" :wrap="false" labelClass="btn btn-info"
                        label="0%" name="agriculturaIzq" value="5" />
                    </template>
                    <template #tip> Sin cultivos o zonas para ganado ni derivaciones para estas actividades</template>
                  </MDBTooltip>
                </MDBCol>
                <MDBCol>
                  <MDBTooltip v-model="tooltip22" direction="bottom">
                    <template #reference>
                      <MDBRadio required v-model="form.agriculturaIzq" :btnCheck="true" :wrap="false" labelClass="btn btn-info"
                      label=" 20%" name="agriculturaIzq" value="3"  />
                    </template>
                    <template #tip> Suelo para uso agrícola y ganadero</template>
                  </MDBTooltip>
                </MDBCol>
                <MDBCol >
                  <MDBTooltip v-model="tooltip23" direction="bottom">
                    <template #reference>
                      <MDBRadio required v-model="form.agriculturaIzq" :btnCheck="true" :wrap="false" labelClass="btn btn-info"
                      label=" 50%" name="agriculturaIzq" value="2" />
                      </template>
                    <template #tip> Suelo para uso agrícola y ganadero</template>
                  </MDBTooltip>
                </MDBCol>
                <MDBCol >
                  <MDBTooltip v-model="tooltip24" direction="bottom">
                    <template #reference>
                      <MDBRadio required v-model="form.agriculturaIzq" :btnCheck="true" :wrap="false" labelClass="btn btn-info"
                      label=">80%" name="agriculturaIzq" value="0" />
                    </template>
                    <template #tip> Uso agrícola y ganadero, con derivaciones para uso doméstico o industrial</template>
                  </MDBTooltip>
                </MDBCol>
              </MDBBtnGroup>
            </MDBCol>
          </MDBRow>
          <MDBRow> 
            <MDBCol >
              <h5> <b>Hidrología</b> </h5>
            </MDBCol>
          </MDBRow> 
          <MDBRow> 
            <MDBCol >
              <h6>Presas <br/> (interrupciones del flujo)</h6> 
            </MDBCol>
            <MDBCol >
            <MDBBtnGroup class="shadow-0">
                <MDBCol >
                  <MDBTooltip v-model="tooltip25" direction="bottom">
                    <template #reference>
                      <MDBRadio required v-model="form.presas"  :btnCheck="true" :wrap="false" labelClass="btn btn-info"
                        label="Ausencia" name="presas" value="10" />
                      </template>
                    <template #tip>No hay presas (ni de gavión o de costales) corriente arriba del río</template>
                  </MDBTooltip>
                </MDBCol>
                <MDBCol >
                  <MDBTooltip v-model="tooltip26" direction="bottom">
                    <template #reference>
                      <MDBRadio required v-model="form.presas" :btnCheck="true" :wrap="false" labelClass="btn btn-info"
                        label="Presencia" name="presas" value="0"  />
                    </template>
                    <template #tip>Hay presas (incluyendo de gavión y de costales) corriente arriba del río</template>
                  </MDBTooltip>
                </MDBCol>
              </MDBBtnGroup>
            </MDBCol>
          </MDBRow>
          <MDBRow> 
            <MDBCol >
              <h6>Regímenes de velocidad/<br/>profundidad (# de regím. presentes)</h6>
            </MDBCol>
            <MDBCol >
              <MDBBtnGroup class="shadow-0">
                <MDBCol >
                  <MDBTooltip v-model="tooltip27" direction="bottom">
                    <template #reference>
                      <MDBRadio required v-model="form.regimenes"  :btnCheck="true" :wrap="false" labelClass="btn btn-info"
                        label="4 reg." name="regimenes" value="10" />
                    </template>
                    <template #tip> Lento-profundo, lento-somero, rápido-profundo y rápido-somero</template>
                  </MDBTooltip>
                </MDBCol>
                <MDBCol>
                  <MDBTooltip v-model="tooltip28" direction="bottom">
                    <template #reference>
                       <MDBRadio required v-model="form.regimenes" :btnCheck="true" :wrap="false" labelClass="btn btn-info"
                         label="3 reg." name="regimenes" value="7"  />
                    </template>
                    <template #tip>Tres regímenes presentes</template>
                  </MDBTooltip>
                </MDBCol>
                <MDBCol >
                  <MDBTooltip v-model="tooltip29" direction="bottom">
                    <template #reference>
                      <MDBRadio required v-model="form.regimenes" :btnCheck="true" :wrap="false" labelClass="btn btn-info"
                        label="2 reg." name="regimenes" value="4" />
                      </template>
                    <template #tip>Dos regímenes presentes</template>
                  </MDBTooltip>
                </MDBCol>
                <MDBCol >
                  <MDBTooltip v-model="tooltip30" direction="bottom">
                    <template #reference>
                     <MDBRadio required v-model="form.regimenes" :btnCheck="true" :wrap="false" labelClass="btn btn-info"
                       label="1 rég." name="regimenes" value="1" />
                    </template>
                    <template #tip> Un régimen presente, usualmente lento-somero </template>
                  </MDBTooltip>
                </MDBCol>
              </MDBBtnGroup>
            </MDBCol>
          </MDBRow>
          <MDBRow> 
            <MDBCol >
              <h6>Alteración del canal <br/> (evidencia de canalización)</h6>
            </MDBCol>
            <MDBCol >
              <MDBBtnGroup class="shadow-0">
                <MDBCol >
                  <MDBTooltip v-model="tooltip31" direction="bottom">
                    <template #reference>
                      <MDBRadio required v-model="form.canal"  :btnCheck="true" :wrap="false" labelClass="btn btn-info"
                        label="Óptimo" name="canal" value="10" />
                    </template>
                    <template #tip>Sin canalización</template>
                  </MDBTooltip>
                </MDBCol>
                <MDBCol>
                  <MDBTooltip v-model="tooltip32" direction="bottom">
                    <template #reference>
                       <MDBRadio required v-model="form.canal" :btnCheck="true" :wrap="false" labelClass="btn btn-info"
                         label="Medio" name="canal" value="7"  />
                    </template>
                    <template #tip>Evidencia de canalización en el pasado</template>
                  </MDBTooltip>
                </MDBCol>
                <MDBCol >
                  <MDBTooltip v-model="tooltip33" direction="bottom">
                    <template #reference>
                      <MDBRadio required v-model="form.canal" :btnCheck="true" :wrap="false" labelClass="btn btn-info"
                        label="Malo " name="canal" value="4" />
                      </template>
                    <template #tip>40-80% del escurrimiento canalizado e interrumpido</template>
                  </MDBTooltip>
                </MDBCol>
                <MDBCol >
                  <MDBTooltip v-model="tooltip34" direction="bottom">
                    <template #reference>
                      <MDBRadio required v-model="form.canal" :btnCheck="true" :wrap="false" labelClass="btn btn-info"
                       label="Pobre" name="canal" value="1" />
                    </template>
                    <template #tip> Banco de la ribera con cemento o gavión, 80% del escurrimiento canalizado, hábitat alterado o ausente</template>
                  </MDBTooltip>
                </MDBCol>
              </MDBBtnGroup>
            </MDBCol>
          </MDBRow>
          <MDBRow> 
            <MDBCol >
              <h6>Estado del canal <br/> (cant. de agua y sustrato expuesto)</h6>
            </MDBCol>
            <MDBCol >
              <MDBBtnGroup class="shadow-0">
                <MDBCol >
                  <MDBTooltip v-model="tooltip35" direction="bottom">
                    <template #reference>
                      <MDBRadio required v-model="form.estado"  :btnCheck="true" :wrap="false" labelClass="btn btn-info"
                        label="Óptimo" name="estado" value="10" />
                    </template>
                    <template #tip>Agua hasta la base de ambos bancos, sustrato expuesto mínimamente</template>
                  </MDBTooltip>
                </MDBCol>
                <MDBCol>
                  <MDBTooltip v-model="tooltip36" direction="bottom">
                    <template #reference>
                       <MDBRadio required v-model="form.estado" :btnCheck="true" :wrap="false" labelClass="btn btn-info"
                         label="Medio" name="estado" value="7"  />
                    </template>
                    <template #tip>El agua llena +75% del canal, o 25% del sustrato está expuesto</template>
                  </MDBTooltip>
                </MDBCol>
                <MDBCol >
                  <MDBTooltip v-model="tooltip37" direction="bottom">
                    <template #reference>
                      <MDBRadio required v-model="form.estado" :btnCheck="true" :wrap="false" labelClass="btn btn-info"
                        label="Malo " name="estado" value="4" />
                      </template>
                    <template #tip>El agua llena 25-75% del canal, o el sustrato está expuesto</template>
                  </MDBTooltip>
                </MDBCol>
                <MDBCol >
                  <MDBTooltip v-model="tooltip38" direction="bottom">
                    <template #reference>
                      <MDBRadio required v-model="form.estado" :btnCheck="true" :wrap="false" labelClass="btn btn-info"
                       label="Pobre" name="estado" value="1" />
                    </template>
                    <template #tip> Muy poca agua en el canal</template>
                  </MDBTooltip>
                </MDBCol>
              </MDBBtnGroup>
            </MDBCol>
          </MDBRow>
          <MDBRow> 
            <MDBCol >
              <h5> <b>Perturbaciones antropogénicas</b> </h5>
            </MDBCol>
          </MDBRow> 
          <MDBRow> 
            <MDBCol >
              <h6>Efluentes por uso <br/> doméstico </h6> 
            </MDBCol>
            <MDBCol >
            <MDBBtnGroup class="shadow-0">
                <MDBCol >
                  <MDBTooltip v-model="tooltip39" direction="bottom">
                    <template #reference>
                      <MDBRadio required v-model="form.efluentes"  :btnCheck="true" :wrap="false" labelClass="btn btn-info"
                        label="Ausencia" name="efluentes" value="10" />
                      </template>
                    <template #tip>Sin descargas directas de residuos al río </template>
                  </MDBTooltip>
                </MDBCol>
                <MDBCol >
                  <MDBTooltip v-model="tooltip40" direction="bottom">
                    <template #reference>
                      <MDBRadio required v-model="form.efluentes" :btnCheck="true" :wrap="false" labelClass="btn btn-info"
                        label="Presencia" name="efluentes" value="1"  />
                    </template>
                    <template #tip>Hay descargas directas de residuos al río </template>
                  </MDBTooltip>
                </MDBCol>
              </MDBBtnGroup>
            </MDBCol>
          </MDBRow>
          <MDBRow> 
            <MDBCol >
              <h6>Desarrollo urbano <br/>(% uso de suelo)</h6>
            </MDBCol>
            <MDBCol >
              <MDBBtnGroup class="shadow-0">
                <MDBCol >
                  <MDBTooltip v-model="tooltip41" direction="bottom">
                    <template #reference>
                      <MDBRadio required v-model="form.urbano"  :btnCheck="true" :wrap="false" labelClass="btn btn-info"
                        label="0%" name="urbano" value="10" />
                    </template>
                    <template #tip>Sin asentamientos humanos, carreteras o deriv. para usos domésticos o industriales.</template>
                  </MDBTooltip>
                </MDBCol>
                <MDBCol>
                  <MDBTooltip v-model="tooltip42" direction="bottom">
                    <template #reference>
                       <MDBRadio required v-model="form.urbano" :btnCheck="true" :wrap="false" labelClass="btn btn-info"
                         label="20%" name="urbano" value="7"  />
                    </template>
                    <template #tip>Suelo para uso humano</template>
                  </MDBTooltip>
                </MDBCol>
                <MDBCol >
                  <MDBTooltip v-model="tooltip43" direction="bottom">
                    <template #reference>
                      <MDBRadio required v-model="form.urbano" :btnCheck="true" :wrap="false" labelClass="btn btn-info"
                        label="50% " name="urbano" value="4" />
                      </template>
                    <template #tip>Suelo para uso humano</template>
                  </MDBTooltip>
                </MDBCol>
                <MDBCol >
                  <MDBTooltip v-model="tooltip44" direction="bottom">
                    <template #reference>
                      <MDBRadio required v-model="form.urbano" :btnCheck="true" :wrap="false" labelClass="btn btn-info"
                       label=">80%" name="urbano" value="1" />
                    </template>
                    <template #tip> Suelo para uso humano</template>
                  </MDBTooltip>
                </MDBCol>
              </MDBBtnGroup>
            </MDBCol>
          </MDBRow>
          <MDBRow> 
            <MDBCol >
              <h6>Desarrollo humano <br/>(actividades alrededor)</h6>
            </MDBCol>
            <MDBCol >
              <MDBBtnGroup class="shadow-0">
                <MDBCol >
                  <MDBTooltip v-model="tooltip45" direction="bottom">
                    <template #reference>
                      <MDBRadio required v-model="form.humano"  :btnCheck="true" :wrap="false" labelClass="btn btn-info"
                        label="Óptimo" name="humano" value="10" />
                    </template>
                    <template #tip>Ausencia de actividades humanas</template>
                  </MDBTooltip>
                </MDBCol>
                <MDBCol>
                  <MDBTooltip v-model="tooltip46" direction="bottom">
                    <template #reference>
                       <MDBRadio required v-model="form.humano" :btnCheck="true" :wrap="false" labelClass="btn btn-info"
                         label="Medio" name="humano" value="7"  />
                    </template>
                    <template #tip>Al menos 1 actividad: ganadera, agrícola o piscícola</template>
                  </MDBTooltip>
                </MDBCol>
                <MDBCol >
                  <MDBTooltip v-model="tooltip47" direction="bottom">
                    <template #reference>
                      <MDBRadio required v-model="form.humano" :btnCheck="true" :wrap="false" labelClass="btn btn-info"
                        label="Malo " name="humano" value="4" />
                      </template>
                    <template #tip>Al menos 3 actividades: ganadera, agrícola, piscícola,doméstica</template>
                  </MDBTooltip>
                </MDBCol>
                <MDBCol >
                  <MDBTooltip v-model="tooltip48" direction="bottom">
                    <template #reference>
                      <MDBRadio required v-model="form.humano" :btnCheck="true" :wrap="false" labelClass="btn btn-info"
                       label="Pobre" name="humano" value="1" />
                    </template>
                    <template #tip>Actividades agrícolas, ganaderas, piscícolas, domésticas</template>
                  </MDBTooltip>
                </MDBCol>
              </MDBBtnGroup>
            </MDBCol>
          </MDBRow>
          <MDBRow> 
            <MDBCol >
              <h6>Contaminación orgánica y <br/> de basura y escombros</h6>
            </MDBCol>
            <MDBCol >
              <MDBBtnGroup class="shadow-0">
                <MDBCol >
                  <MDBTooltip v-model="tooltip49" direction="bottom">
                    <template #reference>
                      <MDBRadio required v-model="form.contaminacion"  :btnCheck="true" :wrap="false" labelClass="btn btn-info"
                        label="<10%" name="contaminacion" value="10" />
                    </template>
                    <template #tip>% Presencia de basura o escombros</template>
                  </MDBTooltip>
                </MDBCol>
                <MDBCol>
                  <MDBTooltip v-model="tooltip50" direction="bottom">
                    <template #reference>
                       <MDBRadio required v-model="form.contaminacion" :btnCheck="true" :wrap="false" labelClass="btn btn-info"
                         label="20-40%" name="contaminacion" value="7"  />
                    </template>
                    <template #tip> % Presencia de basura o escombros</template>
                  </MDBTooltip>
                </MDBCol>
                <MDBCol >
                  <MDBTooltip v-model="tooltip51" direction="bottom">
                    <template #reference>
                      <MDBRadio required v-model="form.contaminacion" :btnCheck="true" :wrap="false" labelClass="btn btn-info"
                        label="50-80% " name="contaminacion" value="4" />
                      </template>
                    <template #tip>% Presencia de basura o escombros</template>
                  </MDBTooltip>
                </MDBCol>
                <MDBCol >
                  <MDBTooltip v-model="tooltip52" direction="bottom">
                    <template #reference>
                      <MDBRadio required v-model="form.contaminacion" :btnCheck="true" :wrap="false" labelClass="btn btn-info"
                       label=">90%" name="contaminacion" value="1" />
                    </template>
                    <template #tip>% Presencia de basura o escombros</template>
                  </MDBTooltip>
                </MDBCol>
              </MDBBtnGroup>
            </MDBCol>
          </MDBRow>
          <MDBRow> 
            <br> <MDBBtn color="primary" type="submit" >Generar</MDBBtn> 
          </MDBRow>
        </MDBRow>
        <!-- Imagen de ejemplo de llenado -->
         <br><h1 class="display-6">Cómo llenar la ficha del protocolo </h1>
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

<script>
  import { MDBListGroup, 
           MDBListGroupItem,
           MDBTabs,
           MDBTabNav,
           MDBTabContent,
           MDBTabItem,
           MDBTabPane,
           MDBTable,
           MDBBtn,MDBRow,
           MDBCol,
           MDBRadio, 
           MDBBtnGroup,
           MDBTooltip
          } from "mdb-vue-ui-kit";
  import { ref } from 'vue';
  import router from '../router'
  import axios from 'axios';
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
      MDBRow,
      MDBCol,
      MDBRadio,
      MDBBtnGroup,
      MDBTooltip
    },
    setup() {
      const checkForm = (e) => {
        e.target.classList.add('was-validated');
      };
      const activeTabId = ref('info-hyqi');
      const tooltip1 = ref(false);
      const tooltip2 = ref(false);
      const tooltip3 = ref(false);
      const tooltip4 = ref(false);
      const tooltip5 = ref(false);
      const tooltip6 = ref(false);
      const tooltip7 = ref(false);
      const tooltip8 = ref(false);
      const tooltip9 = ref(false);
      const tooltip10 = ref(false);
      const tooltip11 = ref(false);
      const tooltip12 = ref(false);
      const tooltip13 = ref(false);
      const tooltip14 = ref(false);
      const tooltip15 = ref(false);
      const tooltip16 = ref(false);
      const tooltip17 = ref(false);
      const tooltip18 = ref(false);
      const tooltip19 = ref(false);
      const tooltip20 = ref(false);
      const tooltip21 = ref(false);
      const tooltip22 = ref(false);
      const tooltip23 = ref(false);
      const tooltip24 = ref(false);
      const tooltip25 = ref(false);
      const tooltip26 = ref(false);
      const tooltip27 = ref(false);
      const tooltip28 = ref(false);
      const tooltip29 = ref(false);
      const tooltip30 = ref(false);
      const tooltip31 = ref(false);
      const tooltip32 = ref(false);
      const tooltip33 = ref(false);
      const tooltip34 = ref(false);
      const tooltip35 = ref(false);
      const tooltip36 = ref(false);
      const tooltip37 = ref(false);
      const tooltip38 = ref(false);
      const tooltip39 = ref(false);
      const tooltip40 = ref(false);
      const tooltip41 = ref(false);
      const tooltip42 = ref(false);
      const tooltip43 = ref(false);
      const tooltip44 = ref(false);
      const tooltip45 = ref(false);
      const tooltip46 = ref(false);
      const tooltip47 = ref(false);
      const tooltip48 = ref(false);
      const tooltip49 = ref(false);
      const tooltip50 = ref(false);
      const tooltip51 = ref(false);
      const tooltip52 = ref(false);
      return {
        activeTabId,
        checkForm,
        tooltip1,
        tooltip2,
        tooltip3,
        tooltip4,
        tooltip5,
        tooltip6,
        tooltip7,
        tooltip8,
        tooltip9,
        tooltip10,
        tooltip11,
        tooltip12,
        tooltip13,
        tooltip14,
        tooltip15,
        tooltip16,
        tooltip17,
        tooltip18,
        tooltip19,
        tooltip20,
        tooltip21,
        tooltip22,
        tooltip23,
        tooltip24,
        tooltip25,
        tooltip26,
        tooltip27,
        tooltip28,
        tooltip29,
        tooltip30,
        tooltip31,
        tooltip32,
        tooltip33,
        tooltip34,
        tooltip35,
        tooltip36,
        tooltip37,
        tooltip38,
        tooltip39,
        tooltip40,
        tooltip41,
        tooltip42,
        tooltip43,
        tooltip44,
        tooltip45,
        tooltip46,
        tooltip47,
        tooltip48,
        tooltip49,
        tooltip50,
        tooltip51,
        tooltip52,
        form: {
          coberturaDer: null,
          coberturaIzq: null,
          estabilidad: null,
          sustrato: null,
          agriculturaDer: null,
          agriculturaIzq: null,
          presas: null,
          regimenes: null,
          canal: null,
          estado: null,
          efluentes: null,
          urbano: null,
          humano: null,
          contaminacion: null
        } 
      };
    },
    methods: {
      navega: function(url) {
        router.push(url).catch(()=>{});
        },
      genera(event){
        event.preventDefault()
        console.log("se llega aqui")
        axios.post(`/api/hyqi/export/excel`, JSON.stringify(this.form)  ,{ 
          headers: {
            'Content-Type': 'application/json',
            Accept: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
          }, responseType : 'blob'}
          ).then(resp => {
            const url = window.URL.createObjectURL(new Blob([resp.data],
              { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" })); //excel file
            const link = document.createElement('a');
            link.href = url;
            link.setAttribute('download', 'ProtocoloHYQI.xlsx'); 
            document.body.appendChild(link);
            link.click();
          });
      }
    }  
  };
</script>
