import numpy as np

import Data_Store
import Power_flow


class Probabilistic_NR_load_flow(Power_flow.NR_class):
    def __init__(self, num_buses, V_array, Y_bus_matrix, PG_array, QG_array,
                 PL_array, QL_array, Q_min_array, Q_max_array, slack_bus_num, bus_type_array,
                 std_dev_array, num_simulations, max_iterations, tolerance):
        self.PG_array, self.QG_array, self.PL_array, self.QL_array, self.iteration_result = None, None, None, None, None
        self.PG_array_init, self.QG_array_init = PG_array, QG_array
        self.QL_array_init, self.PL_array_init = QL_array, PL_array
        self.max_iterations, self.tolerance, self.Y_bus_matrix = max_iterations, tolerance, Y_bus_matrix
        self.V_array_init, self.num_buses = V_array, num_buses
        self.Q_min_array, self.Q_max_array = Q_min_array, Q_max_array
        self.slack_bus_num, self.bus_type_array = slack_bus_num, bus_type_array
        self.num_simulations = num_simulations
        self.std_dev_PL, self.std_dev_QL = std_dev_array[0], std_dev_array[1]
        self.std_dev_PG, self.std_dev_QG = std_dev_array[2], std_dev_array[3]
        self.temp_PL_array_list, self.temp_PG_array_list = [], []

    def run_probabilistic_load_flow(self):
        non_convergence_count, non_convergence_log, all_results, PLF_V_array = 0, [], [], []
        for simulation in range(self.num_simulations):
            temp_PL_array = np.maximum(0, np.random.normal(self.PL_array_init, self.std_dev_PL))
            temp_QL_array = np.random.normal(self.QL_array_init, self.std_dev_QL)
            temp_PG_array = np.maximum(0, np.random.normal(self.PG_array_init, self.std_dev_PG))
            temp_QG_array = np.random.normal(self.QG_array_init, self.std_dev_QG)
            self.temp_PG_array_list.append(temp_PG_array)
            self.temp_PL_array_list.append(temp_PL_array)

            self.PL_array, self.QL_array, self.PG_array, self.QG_array = (
                temp_PL_array, temp_QL_array, temp_PG_array, temp_QG_array)
            super().__init__(self.num_buses, self.V_array_init, self.Y_bus_matrix, temp_PG_array, temp_QG_array,
                             temp_PL_array, temp_QL_array, self.Q_min_array, self.Q_max_array,
                             self.slack_bus_num, self.bus_type_array, self.max_iterations, self.tolerance)
            self.V_array, self.iteration_result = self.NR_solve()
            PLF_V_array.append(self.V_array.copy())

            print(f"Simulation {simulation + 1}: Voltage magnitudes - {np.abs(self.V_array)}")
            print(f"std_dev_PL: {self.std_dev_PL}, temp_PL_array (first few values): {temp_PL_array[:3]}")
            if not self.iteration_result.startswith("Converged"):
                non_convergence_count += 1
                non_convergence_log.append({'simulation': simulation + 1})

        PLF_V_array = np.array(PLF_V_array)
        mags, angles = np.abs(PLF_V_array), np.angle(PLF_V_array)
        Data_Store.central_data_store.update_data('mags', mags)
        Data_Store.central_data_store.update_data('angles', angles)
        average_mags_array = np.mean(mags, axis=0)
        average_ang_array = np.mean(angles, axis=0)

        min_mag_index = np.argmin(average_mags_array)
        max_mag_index = np.argmax(average_mags_array)
        min_mag = average_mags_array[min_mag_index]
        min_angle = average_ang_array[min_mag_index]
        max_mag = average_mags_array[max_mag_index]
        max_angle = average_ang_array[max_mag_index]

        all_results.append({
            "Average Magnitude": average_mags_array, "Average Angle": average_ang_array,
            "Lowest Magnitude": min_mag, "Lowest Angle": min_angle,
            "Maximum Magnitude": max_mag, "Maximum Angle": max_angle
        })
        Data_Store.central_data_store.update_data('all_results', all_results)
        Data_Store.central_data_store.update_data('non_convergence_log', non_convergence_log)
        Data_Store.central_data_store.update_data('temp_PG_array_list', self.temp_PG_array_list)
        Data_Store.central_data_store.update_data('temp_PL_array_list', self.temp_PL_array_list)

        print("All results:", all_results)
        print("Non-convergence log:", non_convergence_log)
        return all_results, non_convergence_log


class Probabilistic_NRFD_load_flow(Power_flow.NRFD_class):
    def __init__(self, num_buses, V_array, Y_bus_matrix, PG_array, QG_array,
                 PL_array, QL_array, Q_min_array, Q_max_array, slack_bus_num, bus_type_array,
                 std_dev_array, num_simulations, max_iterations, tolerance):
        self.PG_array, self.QG_array, self.PL_array, self.QL_array, self.iteration_result = None, None, None, None, None
        self.PG_array_init, self.QG_array_init = PG_array, QG_array
        self.QL_array_init, self.PL_array_init = QL_array, PL_array
        self.max_iterations, self.tolerance, self.Y_bus_matrix = max_iterations, tolerance, Y_bus_matrix
        self.V_array_init, self.num_buses = V_array, num_buses
        self.Q_min_array, self.Q_max_array = Q_min_array, Q_max_array
        self.slack_bus_num, self.bus_type_array = slack_bus_num, bus_type_array
        self.num_simulations = num_simulations
        self.std_dev_PL, self.std_dev_QL = std_dev_array[0], std_dev_array[1]
        self.std_dev_PG, self.std_dev_QG = std_dev_array[2], std_dev_array[3]
        self.temp_PL_array_list, self.temp_PG_array_list = [], []

    def run_probabilistic_load_flow(self):
        non_convergence_count, non_convergence_log, all_results, PLF_V_array = 0, [], [], []
        for simulation in range(self.num_simulations):
            temp_PL_array = np.maximum(0, np.random.normal(self.PL_array_init, self.std_dev_PL))
            temp_QL_array = np.random.normal(self.QL_array_init, self.std_dev_QL)
            temp_PG_array = np.maximum(0, np.random.normal(self.PG_array_init, self.std_dev_PG))
            temp_QG_array = np.random.normal(self.QG_array_init, self.std_dev_QG)
            self.temp_PG_array_list.append(temp_PG_array)
            self.temp_PL_array_list.append(temp_PL_array)

            self.PL_array, self.QL_array, self.PG_array, self.QG_array = (
                temp_PL_array, temp_QL_array, temp_PG_array, temp_QG_array)
            super().__init__(self.num_buses, self.V_array_init, self.Y_bus_matrix, temp_PG_array, temp_QG_array,
                             temp_PL_array, temp_QL_array, self.Q_min_array, self.Q_max_array,
                             self.slack_bus_num, self.bus_type_array, self.max_iterations, self.tolerance)
            self.V_array, self.iteration_result = self.NRFD_solve()
            PLF_V_array.append(self.V_array.copy())

            print(f"Simulation {simulation + 1}: Voltage magnitudes - {np.abs(self.V_array)}")
            print(f"std_dev_PL: {self.std_dev_PL}, temp_PL_array (first few values): {temp_PL_array[:3]}")
            if not self.iteration_result.startswith("Converged"):
                non_convergence_count += 1
                non_convergence_log.append({'simulation': simulation + 1})

        PLF_V_array = np.array(PLF_V_array)
        mags, angles = np.abs(PLF_V_array), np.angle(PLF_V_array)
        Data_Store.central_data_store.update_data('mags', mags)
        Data_Store.central_data_store.update_data('angles', angles)
        average_mags_array = np.mean(mags, axis=0)
        average_ang_array = np.mean(angles, axis=0)

        min_mag_index = np.argmin(average_mags_array)
        max_mag_index = np.argmax(average_mags_array)
        min_mag = average_mags_array[min_mag_index]
        min_angle = average_ang_array[min_mag_index]
        max_mag = average_mags_array[max_mag_index]
        max_angle = average_ang_array[max_mag_index]

        all_results.append({
            "Average Magnitude": average_mags_array, "Average Angle": average_ang_array,
            "Lowest Magnitude": min_mag, "Lowest Angle": min_angle,
            "Maximum Magnitude": max_mag, "Maximum Angle": max_angle
        })
        Data_Store.central_data_store.update_data('all_results', all_results)
        Data_Store.central_data_store.update_data('non_convergence_log', non_convergence_log)
        Data_Store.central_data_store.update_data('temp_PG_array_list', self.temp_PG_array_list)
        Data_Store.central_data_store.update_data('temp_PL_array_list', self.temp_PL_array_list)

        print("All results:", all_results)
        print("Non-convergence log:", non_convergence_log)
        return all_results, non_convergence_log
