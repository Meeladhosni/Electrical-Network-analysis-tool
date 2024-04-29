import matplotlib
import matplotlib.pyplot as plt
import networkx as nx
import numpy as np
import os
import pandas as pd
import seaborn as sns
import Data_Store
matplotlib.use('TkAgg')


class Load_gen_class:
    def __init__(self):
        self.num_buses = Data_Store.central_data_store.get_data('num_buses')
        self.PG_array = Data_Store.central_data_store.get_data('PG_output_array')
        self.PL_array = Data_Store.central_data_store.get_data('PL_array')
        self.PG_sub_lists = [Data_Store.central_data_store.get_data(f'PG_sub_{i}')
                             for i in range(1, self.num_buses + 1)]
        self.bar_width = 0.1
        self.offset = 0.05

    def plot_Load_gen(self):
        plt.figure(figsize=(12, 6))
        # Calculate positions
        indices = np.arange(self.num_buses)
        load_positions = indices - 0.2
        gen_positions = indices + 0.2

        # Main Load and Generation
        plt.bar(load_positions, self.PL_array, width=self.bar_width,
                color='blue', edgecolor='black', label='Load per Bus')
        plt.bar(gen_positions, self.PG_array, width=self.bar_width,
                color='green', edgecolor='black', label='Generation per Bus')

        # Sub-generation
        sub_gen_positions = indices + 0.2 + self.bar_width
        for idx, pg_subs in enumerate(self.PG_sub_lists):
            if pg_subs:
                valid_pg_subs = [val for val in pg_subs if val > 0]
                if len(valid_pg_subs) > 1:  # Only plot if there's more than one non-zero value
                    sub_bar_positions = sub_gen_positions[idx] + np.arange(len(valid_pg_subs)) * self.offset
                    plt.bar(sub_bar_positions, valid_pg_subs, width=self.bar_width, color='lightgreen',
                            edgecolor='black', label=f'Sub-Generation Bus {idx + 1}' if idx == 0 else "")

        # Total Load and Generation
        total_load = sum(self.PL_array)
        total_gen = sum(self.PG_array)
        plt.bar(-0.8, total_load, width=self.bar_width, color='navy', edgecolor='black', label='Total Load')
        plt.bar(-0.6, total_gen, width=self.bar_width, color='darkgreen', edgecolor='black', label='Total Generation')

        # Adding labels, titles, and customizing x-ticks
        plt.xlabel('Bus Number')
        plt.ylabel('Power (p.u.)')
        plt.title('Load and Generation at Each Bus')
        plt.xticks(indices, range(1, self.num_buses + 1))
        plt.legend()
        plt.grid(True)

    @staticmethod
    def plt_show():
        plt.tight_layout()
        plt.show()

    @staticmethod
    def plt_save():
        plot_file_path = os.path.join('config', 'Load_gen_plot.png')
        plt.savefig(plot_file_path)
        plt.close()

class V_profile_class:
    def __init__(self):
        self.num_buses = Data_Store.central_data_store.get_data('num_buses')
        self.V_mag_output_array = Data_Store.central_data_store.get_data('V_mag_output_array')
        self.V_ang_output_array = Data_Store.central_data_store.get_data('V_ang_output_array')

    def plot_V_profile(self):
        bus_numbers = np.arange(1, self.num_buses + 1)

        # Plotting Voltage Profile Plot
        plt.figure(figsize=(12, 5))
        plt.subplot(1, 2, 1)
        plt.plot(bus_numbers, self.V_mag_output_array, marker='o', linestyle='-', color='b')
        plt.xlabel('Bus Number')
        plt.ylabel('Voltage Magnitude (p.u.)')
        plt.title('Voltage Profile Plot')
        plt.grid(True)

        # Adding a horizontal line to indicate nominal voltage level
        plt.axhline(y=1.0, color='r', linestyle='--')
        plt.xticks(bus_numbers)

        # Plotting Phase Angle Plot
        plt.subplot(1, 2, 2)
        plt.plot(bus_numbers, self.V_ang_output_array, marker='x', linestyle='-', color='g')
        plt.xlabel('Bus Number')
        plt.ylabel('Phase Angle (degrees)')
        plt.title('Phase Angle Plot')
        plt.grid(True)
        plt.xticks(bus_numbers)

    @staticmethod
    def plt_show():
        plt.tight_layout()
        plt.show()

    @staticmethod
    def plt_save():
        plot_file_path = os.path.join('config', 'V_profile_plot.png')
        plt.savefig(plot_file_path)
        plt.close()

class PF_class:
    def __init__(self):
        self.num_buses = Data_Store.central_data_store.get_data('num_buses')
        self.PG_array = Data_Store.central_data_store.get_data('PG_output_array')
        self.QG_array = Data_Store.central_data_store.get_data('QG_output_array')
        self.PL_array = Data_Store.central_data_store.get_data('PL_array')
        self.QL_array = Data_Store.central_data_store.get_data('QL_array')

    def PF_plot(self):
        # Initialize lists for buses and power factors
        base_buses = [f'Bus {i+1}' for i in range(self.num_buses)]  # Assuming bus numbering starts at 1
        buses = ["Network PF"]
        buses.extend(base_buses)
        P_net = np.sum(self.PG_array) - np.sum(self.PL_array)
        Q_net = np.sum(self.QG_array) - np.sum(self.QL_array)
        S_net = np.sqrt(P_net ** 2 + Q_net ** 2)
        PF = P_net / S_net if S_net != 0 else 0  # Avoid division by zero
        power_factors = [PF]

        # Calculate power factor for each bus
        for i in range(self.num_buses):
            PG = self.PG_array[i]
            QG = self.QG_array[i]
            PL = self.PL_array[i]
            QL = self.QL_array[i]
            P_net = PG - PL
            Q_net = QG - QL
            S_net = np.sqrt(P_net ** 2 + Q_net ** 2)
            PF = P_net / S_net if S_net != 0 else 0  # Avoid division by zero
            power_factors.append(np.abs(PF))
        # Create bar chart for power factors
        fig, ax = plt.subplots(figsize=(8, 4))
        ax.bar(buses, power_factors, color='skyblue', alpha=0.7)

        # Add line for desired power factor level
        desired_pf_line = 0.95
        ax.axhline(y=desired_pf_line, color='r', linestyle='--')

        # Set labels and title
        ax.set_ylabel('Power Factor')
        ax.set_xlabel('Bus')
        ax.set_title('Power Factor at Various Buses')
        ax.set_ylim([0, 1])  # Power factor values range from 0 to 1

        # Annotate desired power factor line
        ax.text(self.num_buses - 1, desired_pf_line, 'Desired Power Factor', va='bottom', ha='right', color='red')

    @staticmethod
    def plt_show():
        plt.tight_layout()
        plt.show()

    @staticmethod
    def plt_save():
        plot_file_path = os.path.join('config', 'PF_plot.png')
        plt.savefig(plot_file_path)
        plt.close()

class PowerNetwork:
    def __init__(self, case):
        self.G = nx.Graph()
        self.num_buses = Data_Store.central_data_store.get_data('num_buses')
        self.add_bus()

        if case:
            self.S_ij_array = Data_Store.central_data_store.get_data('S_ij_array')
            self.V_mag_output_array = Data_Store.central_data_store.get_data('V_mag_output_array')
            self.V_ang_output_array = Data_Store.central_data_store.get_data('V_ang_output_array')
            self.connect_buses(case)
            self.assign_voltages()
        else:
            self.Y_bus_matrix = Data_Store.central_data_store.get_data('Y_bus_matrix')
            self.connect_buses(case)

    def add_bus(self):
        for i in range(1, self.num_buses + 1):
            self.G.add_node(f"Bus {i}")

    def connect_buses(self, case):
        if case:
            S_ij_array_mag = np.abs(self.S_ij_array)
            S_ij_array_ang = np.angle(self.S_ij_array, deg=True)
            for i in range(self.num_buses):
                for j in range(self.num_buses):
                    if self.S_ij_array[i][j].real != 0 and self.S_ij_array[i][j].imag != 0:  # Check if not zero
                        # Access each element with S_ij_array[i][j]
                        self.G.add_edge(f"Bus {i + 1}", f"Bus {j + 1}",
                                        label=(f"{np.round(S_ij_array_mag[i][j], 2)}"
                                               f"∠{np.round(S_ij_array_ang[i][j], 2)}°"))
        else:
            for i in range(self.num_buses):
                for j in range(self.num_buses):
                    if self.Y_bus_matrix[i][j].real != 0 and self.Y_bus_matrix[i][j].imag != 0:  # Check if not zero
                        self.G.add_edge(f"Bus {i + 1}", f"Bus {j + 1}")

    def assign_voltages(self):
        for i in range(self.num_buses):
            bus_name = f"Bus {i + 1}"
            if bus_name in self.G:
                self.G.nodes[bus_name]['voltage'] = (f"{np.round(self.V_mag_output_array[i], 2)}"
                                                     f"∠{np.round(self.V_ang_output_array[i], 2)}°")
            else:
                pass

    def visualize_network(self):
        # Turn on interactive mode so plt.show() is non-blocking
        plt.ion()

        # Check if a figure window is already open
        if plt.fignum_exists(1):
            # Close it if it's open to avoid overlaying figures
            plt.close(1)

        # Create a new figure window
        plt.figure(1)

        # Layout the network once here
        pos = nx.spring_layout(self.G)

        # Offset the label positions to below the node by adjusting y value
        label_pos = {key: [value[0], value[1] - 0.1] for key, value in pos.items()}

        # Draw the network with the nodes labeled by the bus names
        nx.draw(self.G, pos, with_labels=True, node_size=3000, node_color='lightblue', font_size=12)

        # Draw the voltages below the buses using the offset label positions
        labels = {node: f"{data.get('voltage', '')}" for node, data in self.G.nodes(data=True)}
        nx.draw_networkx_labels(self.G, label_pos, labels=labels, font_size=10)

        # Draw the edge labels if they exist
        edge_labels = nx.get_edge_attributes(self.G, 'label')
        if edge_labels:
            nx.draw_networkx_edge_labels(self.G, pos, edge_labels=edge_labels)

        # Display the plot without blocking the rest of the code
        plt.show(block=False)
        plot_file_path = os.path.join('config', 'network_plot.png')
        # Save the currently active figure
        plt.savefig(plot_file_path)

class Histogram_class:
    def __init__(self):
        self.num_simulations = Data_Store.central_data_store.get_data('num_simulations')
        self.num_buses = Data_Store.central_data_store.get_data('num_buses')
        self.bus_ids = []
        self.mags = Data_Store.central_data_store.get_data('mags')
        self.angles = Data_Store.central_data_store.get_data('angles')
        self.slack_bus_index = Data_Store.central_data_store.get_data('slack_bus_num') - 1
        self.mags[:, self.slack_bus_index] = 0

    def Histogram_plot(self, bus_index):
        plt.figure(figsize=(12, 6))
        # Histogram for Voltage Magnitudes
        plt.subplot(1, 2, 1)
        if bus_index is not None:
            # Plot only the specified bus
            plt.hist(np.round(self.mags[:, bus_index], 3), bins=20, alpha=0.5, label=f'Bus {bus_index + 1}')
            plt.title(f'Distribution of Voltage Magnitudes for Bus {bus_index + 1}')
            print(f"Bus {bus_index + 1} magnitude data:", self.mags[:, bus_index])
        else:
            for i in range(self.num_buses):
                if i == self.slack_bus_index:
                    continue
                plt.hist(self.mags[:, i], bins=20, alpha=0.5, label=f'Bus {i + 1}', density=True)
            plt.title('Distribution of Voltage Magnitudes for All Buses Except Slack')
        plt.xlabel('Voltage Magnitude (p.u.)')
        plt.ylabel('Frequency')
        plt.legend()

        # Histogram for Voltage Angles
        plt.subplot(1, 2, 2)
        if bus_index is not None:
            # Plot only the specified bus
            plt.hist(self.angles[:, bus_index], bins=20, alpha=0.5, label=f'Bus {bus_index + 1}')
            plt.title(f'Distribution of Voltage Angles for Bus {bus_index + 1}')

        else:
            for i in range(self.num_buses):
                if i == self.slack_bus_index:
                    continue
                plt.hist(self.angles[:, i], bins=20, alpha=0.5, label=f'Bus {i + 1}', density=True)
            plt.title('Distribution of Voltage Angles for All Buses Except Slack')
        plt.xlabel('Voltage Angle (degrees)')
        plt.ylabel('Frequency')
        plt.legend()

        plt.tight_layout()
        plt.show()

class Correlation_class:
    def __init__(self, bus_index):
        self.num_buses = Data_Store.central_data_store.get_data('num_buses')
        PG_variations = Data_Store.central_data_store.get_data('temp_PG_array_list')
        PL_variations = Data_Store.central_data_store.get_data('temp_PL_array_list')
        mags = Data_Store.central_data_store.get_data('mags').flatten()
        if bus_index is not None:
            PG = np.array(PG_variations)
            PL = np.array(PL_variations)

            self.PG_variations = PG[:, bus_index]
            self.PL_variations = PL[:, bus_index]
            self.mags = np.round(mags[bus_index::self.num_buses], 3)
            print("PG", self.PG_variations)
            print("mags", self.mags)

        else:
            self.PG_variations = np.concatenate(PG_variations)
            self.PL_variations = np.concatenate(PL_variations)
            self.mags = mags
        self.data = pd.DataFrame({
                'Load Variation': self.PL_variations,
                'Generation Variation': self.PG_variations,
                'Voltage Magnitude': self.mags
        })

    def Load_Variation(self, bus_index):
        # Plot for Load Variation vs. Voltage Magnitude
        plot_1 = sns.jointplot(x='Load Variation', y='Voltage Magnitude', data=self.data, kind='reg', color='b')
        if bus_index:
            plot_1.fig.suptitle(f'Load Variation vs. Voltage mag [Bus {bus_index + 1}')
        else:
            plot_1.fig.suptitle('Load Variation vs. Voltage Magnitude')
        plot_1.fig.tight_layout()
        plot_1.fig.subplots_adjust(top=0.95)  # Adjust the top margin
        plt.show()

    def Generation_Variation(self, bus_index):
        # Plot for Generation Variation vs. Voltage Magnitude
        plot_2 = sns.jointplot(x='Generation Variation', y='Voltage Magnitude', data=self.data, kind='reg', color='g')
        if bus_index:
            plot_2.fig.suptitle(f'Generation Variation vs. Voltage mag [Bus {bus_index + 1}]')
        else:
            plot_2.fig.suptitle('Generation Variation vs. Voltage Magnitude')
        plot_2.fig.tight_layout()
        plot_2.fig.subplots_adjust(top=0.95)  # Adjust the top margin
        plt.show()

    def Load_Gen_Variation(self, bus_index):
        fig = plt.figure()
        ax = fig.add_subplot(111, projection='3d')

        # Scatter plot
        ax.scatter(self.data['Load Variation'], self.data['Generation Variation'],
                       self.data['Voltage Magnitude'], c='b', marker='o')

        # Labels and titles
        ax.set_xlabel('Load Variation')
        ax.set_ylabel('Generation Variation')
        ax.set_zlabel('Voltage Magnitude')
        if bus_index:
            ax.set_title(f'Load and Generation Variation vs. Voltage mag [Bus {bus_index + 1}]')
        else:
            ax.set_title(f'Load and Generation Variation vs. Voltage Magnitude')
        plt.show()
