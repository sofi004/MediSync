import React from 'react';
import { Text, View, StyleSheet, ScrollView, Image} from 'react-native';

export default function HowToUseScreen() {
  return (
    <ScrollView style={styles.container}>
      <Image
        source={require('../../assets/images/excel-example.png')}
        style={styles.image}
      />
      <Text style={styles.heading}>Excel File Structure:</Text>
      <Text style={styles.paragraph}>To ensure the application functions correctly, the Excel file you upload must adhere to the following structure:</Text>

      <View style={styles.listItem}>
        <Text style={styles.bold}>Column 1:</Text>
        <Text style={styles.whiteText}> List all hospital rooms relevant to the managed medical team.</Text>
      </View>

      <View style={styles.listItem}>
        <Text style={styles.bold}>Column 2:</Text>
        <Text style={styles.whiteText}> Include the names of all doctors in the team.</Text>
      </View>

      <View style={styles.listItem}>
        <Text style={styles.bold}>Column 3:</Text>
        <Text style={styles.whiteText}> Contain the names of all possible activities that can occur in the hospital.</Text>
      </View>

      <View style={styles.listItem}>
        <Text style={styles.bold}>Column 4:</Text>
        <Text style={styles.whiteText}> Specify each doctor's absences. For each doctor, provide one or more date ranges indicating their unavailability.</Text>
      </View>

      <Text style={styles.heading}>Recurring Weekly Schedules ("Week Templates"):</Text>
      <Text style={styles.paragraph}>Following the initial columns, you need to define the recurring weekly schedules. Each week template is structured as follows:</Text>

      <View style={styles.listItem}>
        <Text style={styles.bold}>Morning and Afternoon Schedules:</Text>
        <Text style={styles.whiteText}> Each week template comprises two sections: a morning schedule and an afternoon schedule.</Text>
      </View>

      <View style={styles.listItem}>
        <Text style={styles.bold}>Time Slots (Rows):</Text>
        <Text style={styles.whiteText}> Each row within a schedule (morning or afternoon) represents a specific time slot.</Text>
      </View>

      <View style={styles.listItem}>
        <Text style={styles.bold}>Information per Time Slot (Columns):</Text>
        <Text style={styles.whiteText}> Each row must contain the following information in separate columns:</Text>
        <View style={styles.nestedListItem}>
          <Text style={styles.whiteText}>- Doctor's Name</Text>
          <Text style={styles.whiteText}>- Assigned Room</Text>
          <Text style={styles.whiteText}>- Scheduled Activity</Text>
        </View>
      </View>

      <View style={styles.listItem}>
        <Text style={styles.bold}>Structure of a Week Template:</Text>
        <Text style={styles.whiteText}> A week template consists of a consecutive group of rows for the morning schedule, immediately followed by a consecutive group of rows for the afternoon schedule.</Text>
      </View>

      <View style={styles.listItem}>
        <Text style={styles.bold}>New Week Template Indication:</Text>
        <Text style={styles.whiteText}> A new set of morning schedule rows appearing after an afternoon schedule section signifies the start of a new week template.</Text>
        <Text></Text>
      </View>
    </ScrollView>
  );
}

const styles = StyleSheet.create({
  container: {
    flex: 1,
    backgroundColor: '#25292e',
    padding: 20,
  },
  title: {
    fontSize: 24,
    fontWeight: 'bold',
    marginBottom: 20,
    color: '#fff',
  },
  heading: {
    fontSize: 18,
    fontWeight: 'bold',
    marginTop: 15,
    marginBottom: 10,
    color: '#ffd33d',
  },
  paragraph: {
    fontSize: 16,
    marginBottom: 15,
    color: '#fff',
    lineHeight: 24,
  },
  listItem: {
    marginBottom: 10,
  },
  bold: {
    fontWeight: 'bold',
    color: '#fff',
  },
  whiteText: {
    color: '#fff',
  },
  nestedListItem: {
    marginLeft: 15,
  },
  image: {
    width: 350, // Ajuste a largura conforme necessário
    height: 250, // Ajuste a altura conforme necessário
    resizeMode: 'contain', // Ou 'cover', 'stretch', 'repeat', 'center'
    marginBottom: 20,
  },
});