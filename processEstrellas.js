const fs = require('fs');
const csv = require('csv-parser');
const docx = require('docx');

const { Document, Packer, Paragraph, TextRun } = docx;

// Initialize the 'estrellas' list
const estrellas = [];

// Esta funcion limpia el campo del nombre
function limpiarNombres(text) {
  // Split the string into an array of elements
  const elementArray = text.split(';#');

  // Use the filter method to remove elements that are numbers only
  const filteredArray = elementArray.filter(
    (element) => !/^[0-9]+$/.test(element)
  );

  // Use the map method and a regular expression to remove the unwanted characters from each element
  const cleanedElements = filteredArray.map((element) =>
    element.replace(/\s*\(.*\)/, '')
  );

  const capitalizedElements = cleanedElements.map((element) =>
    element
      .toLowerCase()
      .split(' ')
      .map((word) => word.charAt(0).toUpperCase() + word.slice(1))
      .join(' ')
  );

  // Join the array of elements into a single string separated by semicolons
  return capitalizedElements.join('; ');
}

// Esta funcion limpia el campo del cargo
function limpiarCargo(cargo) {
  // Pasamos el cargo a capitalizacion normal si está en todo mayúsculas
  // Reemplazamos los saltos de línea por el separador "; "
  cargo = cargo.replace(/\n/g, ' / ');

  // Si el nombre está en mayúsculas, lo convertimos a capitalizacion normal
  cargo = cargo.toLowerCase().replace(/\b[a-z]/g, function (letra) {
    return letra.toUpperCase();
  });

  return cargo;
}

function limpiarEstacion(estacion) {
  // Check if the input has parentheses
  if (estacion.match(/\(.+\)/)) {
    // Extraemos el codigo de la estacion entre paréntesis
    estacion = estacion.match(/\((.+)\)/)[1];
  }

  return estacion;
}

function filterEstrellas(estrellas) {
  // Get the current date and time
  const now = new Date();

  // Get the last Wednesday at 9am

  const day = now.getDay();
  let daysToMostRecentWednesday = day === 3 ? 0 : 3 - day;
  if (daysToMostRecentWednesday < 0) {
    daysToMostRecentWednesday += 7;
  }

  const lastWednesday9am = new Date(
    now.getFullYear(),
    now.getMonth(),
    now.getDate()
  );
  lastWednesday9am.setHours(9, 0, 0, 0);

  lastWednesday9am.setDate(lastWednesday9am - daysToMostRecentWednesday);

  // Get the date 7 days prior to the last Wednesday at 9am
  const startDate = new Date(lastWednesday9am);
  const endDate = new Date(lastWednesday9am);
  console.log(startDate);
  startDate.setDate(endDate.getDate() - 7);

  // Filter the estrellas to only show those with a date greater than or equal to the start date
  const filteredEstrellas = estrellas.filter(
    (estrella) =>
      new Date(estrella.date) >= startDate && new Date(estrella.date) <= endDate
  );

  return filteredEstrellas;
}

function groupEstrellasByArea(estrellas) {
  // Use the reduce method to create a new object with keys representing the recipient_areas and values representing arrays of the estrellas objects with that recipient_area
  const groupedEstrellas = estrellas.reduce((groups, estrella) => {
    let recipientArea = '';

    if (estrella.recipient_areas.length > 1) {
      recipientArea = 'Trabajamos en conjunto';
    } else if (estrella.passenger_name_and_station) {
      recipientArea =
        'Anticipamos y superamos las expectativas de nuestros clientes';
    } else {
      recipientArea = estrella.recipient_areas[0];
    }

    // If the group for this recipient_area does not exist yet, create it
    if (!groups[recipientArea]) {
      groups[recipientArea] = [];
    }

    // Add the estrella object to the group for its recipient_area
    groups[recipientArea].push(estrella);

    return groups;
  }, {});

  return groupedEstrellas;
}

function estrellasToText(group) {
  let text = '';
  for (let key in group) {
    text += '\n\n' + key + '\n\n';
    group[key].forEach((estrella) => {
      const giver =
        estrella.passenger_name_and_station === ''
          ? estrella.giver_names +
            ', ' +
            estrella.giver_positions +
            ', ' +
            estrella.giver_stations
          : estrella.passenger_name_and_station;
      text +=
        estrella.recipient_names +
        '\n' +
        estrella.recipient_positions +
        ', ' +
        estrella.recipient_stations +
        '\n' +
        estrella.content +
        '\n' +
        giver +
        '\n\n';
    });
  }
  return text;
}

function formatStar(estrella) {
  const giver =
    estrella.passenger_name_and_station === ''
      ? estrella.giver_names + ' - '
      : estrella.passenger_name_and_station;

  const positionAndStation =
    estrella.passenger_name_and_station === ''
      ? estrella.giver_positions + ', ' + estrella.giver_stations
      : '';
  const formattedStar = [
    new Paragraph({
      children: [
        new TextRun({
          text: estrella.recipient_names,
          bold: true,
          color: '#0360AC',
          size: 22,
          font: 'Arial',
          break: 2,
        }),
        new TextRun({
          text:
            estrella.recipient_positions + ', ' + estrella.recipient_stations,
          italics: true,
          color: '#0360AC',
          size: 20,
          font: 'Arial',
          break: 1,
        }),
        new TextRun({
          text: estrella.content,
          color: '#737373',
          size: 20,
          font: 'Arial',
          break: 1,
        }),
        new TextRun({
          text: giver,
          bold: true,
          color: '#9B7615',
          size: 20,
          font: 'Arial',
          break: 1,
        }),
        new TextRun({
          text: positionAndStation,
          italics: true,
          color: '#9B7615',
          size: 20,
          font: 'Arial',
        }),
      ],
    }),
  ];
  return formattedStar;
}

function createSectionTitle(title) {
  const sectionTitle = new Paragraph({
    children: [
      new TextRun({
        text: title.toString(),
        bold: true,
        size: 28,
        font: 'Arial',
        break: 2,
      }),
    ],
  });
  return sectionTitle;
}

function estrellasToDoc(group) {
  console.log('Creating doc...');
  const formattedStars = [];
  for (let key in group) {
    formattedStars.push(createSectionTitle(key));
    group[key].forEach((estrella) => {
      const formattedStar = formatStar(estrella);
      formattedStars.push(...formattedStar);
    });
  }
  const doc = new Document({
    creator: 'Beatriz Gonzalez',
    title: 'Estrellas de la Semana',
    styles: {
      paragraphStyles: [
        {
          paragraph: {
            spacing: {
              line: 276,
              after: 300,
              before: 200,
            },
          },
        },
      ],
    },
    sections: [
      {
        properties: {},
        children: formattedStars,
      },
    ],
  });
  return doc;
}

function processCSVOutput(output) {
  const filteredEstrellas = filterEstrellas(output);
  const groupedEstrellas = groupEstrellasByArea(output);
  const doc = estrellasToDoc(groupedEstrellas);
  const text = estrellasToText(groupedEstrellas);
  
  // Save text output locally if needed
  fs.writeFileSync('estrellas.txt', text);

  // Return the document object for further processing in index.js
  return { doc, text };
}

module.exports = {
  processCSVOutput
};
